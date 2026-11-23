# -*- coding: utf-8 -*-
"""Mico 选品工作台 —— 服务端 + 渲染（分页版）"""
import os
import re
import json
import html
import shutil
import datetime
import traceback
import threading
import time
import urllib.request
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import parse_qs, quote, unquote, urlparse

import openpyxl
from openpyxl.drawing.image import Image as XLImage
import similar
import supabase_store as sb
from openpyxl.styles import Alignment, Border, Side, Font

BASE = os.path.dirname(os.path.abspath(__file__))

# ---- 云部署配置（先定义，供下方路径使用）----
# 数据目录：云上通常挂载持久卷，指向持久化路径
DATA_DIR = os.environ.get("DATA_DIR", BASE)
PORT = int(os.environ.get("PORT", "8000"))
# 云上必须监听 0.0.0.0 才能被外部访问
HOST = os.environ.get("HOST", "127.0.0.1")
# 设 APP_PASSWORD 即启用登录；留空则免登录（仅本机使用）
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")
USERS = [u.strip() for u in os.environ.get("APP_USERS", "同事A,同事B").split(",") if u.strip()]
PER_PAGE = 96

# ---- 文件路径 ----
def _find_inq_file():
    """优先用精简版（无内嵌图，云部署省空间），否则用完整版。"""
    lite = os.path.join(BASE, "已询盘产品_四列表_精简版.xlsx")
    full = os.path.join(BASE, "已询盘产品_四列表.xlsx")
    return lite if os.path.exists(lite) else full


INQ_FILE = _find_inq_file()
DATA_FILE = os.path.join(DATA_DIR, "_data.json")
BACKUP_DIR = os.path.join(DATA_DIR, "备份")
IMG_MAP_FILE = os.path.join(BASE, "_imgmap.json")


os.makedirs(DATA_DIR, exist_ok=True)


def norm(u):
    return (str(u) or "").split("#")[0].split("?")[0].rstrip("/").lower()


# ================= 数据层 =================
_mem_cache = {"data": None, "ts": 0.0}
_CACHE_TTL = 8.0     # 秒，多人协作时兼顾实时性与性能


def invalidate_cache():
    _mem_cache["data"] = None
    _mem_cache["ts"] = 0.0


def load_data():
    """优先从 Supabase 读（带 8 秒内存缓存），失败或未配置则读本地 JSON。"""
    if sb.enabled():
        now = time.time()
        if _mem_cache["data"] is not None and (now - _mem_cache["ts"]) < _CACHE_TTL:
            return _mem_cache["data"]
        try:
            prods = sb.get_products()
            today = datetime.date.today().isoformat()
            # 待采集配对是持续处理的清单，不按日期过滤（旧日期未完成的也要看到）
            pairs = sb.get_pairs()
            news = sb.get_news(today)
            if not news:
                news = sb.get_news()
            data = {"products": prods, "pairs": pairs, "news": news,
                    "pairs_date": today, "updated": "Supabase", "source": "supabase"}
            _mem_cache["data"] = data
            _mem_cache["ts"] = now
            return data
        except Exception as e:
            print("[Supabase读取失败，降级本地] %s" % e)
    if os.path.exists(DATA_FILE):
        try:
            return json.load(open(DATA_FILE, encoding="utf-8"))
        except Exception:
            pass
    return {"products": [], "pairs": [], "news": [], "updated": ""}


_save_lock = threading.Lock()


def save_data(d):
    """写本地缓存（Supabase 由各业务函数单独写）。"""
    d["updated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _save_lock:
        tmp = DATA_FILE + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(d, f, ensure_ascii=False, indent=1)
            try:
                os.replace(tmp, DATA_FILE)
            except Exception:
                time.sleep(0.15)
                try:
                    os.replace(tmp, DATA_FILE)
                except Exception:
                    pass
        except Exception as e:
            print("[本地缓存写入失败] %s" % e)


def load_imgmap():
    if os.path.exists(IMG_MAP_FILE):
        try:
            return json.load(open(IMG_MAP_FILE, encoding="utf-8"))
        except Exception:
            pass
    return {}


def backup_file(path):
    os.makedirs(BACKUP_DIR, exist_ok=True)
    if not os.path.exists(path):
        return None
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    name = os.path.splitext(os.path.basename(path))[0]
    dst = os.path.join(BACKUP_DIR, "%s_%s.xlsx" % (name, ts))
    shutil.copy2(path, dst)
    return os.path.basename(dst)


def sync_from_excel():
    if not os.path.exists(INQ_FILE):
        return {"ok": False, "msg": "找不到 已询盘产品_四列表.xlsx"}
    imgmap = load_imgmap()
    wi = openpyxl.load_workbook(INQ_FILE).active
    products = []
    for r in range(4, wi.max_row + 1):
        nm = wi.cell(row=r, column=2).value
        if not nm:
            continue
        names = [x for x in str(nm).split("\n") if x.strip()]
        url = str(wi.cell(row=r, column=3).value or "")
        iq = str(wi.cell(row=r, column=4).value or "").strip()
        spu_raw = re.sub(r"\s+", " ", str(wi.cell(row=r, column=5).value or "")).strip()
        spus = [s for s in re.split(r"[,\s/;]+", spu_raw) if s]
        u1688 = sb.uniq_urls([
            x.strip() for x in str(wi.cell(row=r, column=6).value or "").split("\n")
            if x.strip().startswith("http")])
        m = re.search(r"INQ-(\d+)", iq)
        products.append({
            "row": r, "names": names, "url": url, "inq": iq,
            "batch": m.group(1)[-4:] if m else "",
            "spu": spu_raw, "spus": spus, "multi_spu": len(spus) > 1,
            "u1688": u1688, "img": imgmap.get(norm(url), ""),
        })
    d = load_data()
    d["products"] = products
    save_data(d)
    return {"ok": True, "count": len(products)}


# ================= 订单处理 =================
def read_order(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    oc = {str(ws.cell(row=1, column=i).value).strip(): i
          for i in range(1, ws.max_column + 1) if ws.cell(row=1, column=i).value}
    for n in ["客户产品链接", "商品名称", "SKU"]:
        if n not in oc:
            raise ValueError("订单文件缺少列：%s" % n)
    lc, nc, sc = oc["客户产品链接"], oc["商品名称"], oc["SKU"]
    ic = oc.get("商品图片")
    prods = {}
    for r in range(2, ws.max_row + 1):
        u = ws.cell(row=r, column=lc).value
        if not u:
            continue
        k = norm(u)
        p = prods.setdefault(k, {"url": str(u), "names": [], "skus": [], "img": ""})
        nm = ws.cell(row=r, column=nc).value
        if nm and str(nm) not in p["names"]:
            p["names"].append(str(nm))
        sku = ws.cell(row=r, column=sc).value
        if sku and str(sku).strip() and str(sku).strip() not in p["skus"]:
            p["skus"].append(str(sku).strip())
        if ic and not p["img"]:
            g = ws.cell(row=r, column=ic).value
            if g and str(g).startswith("http"):
                p["img"] = str(g)
    return prods


def process_order(path):
    """导入订单：合并更新「待采集配对」和「未询盘新品」。
    - 已存在的记录：更新产品信息（名称/图片/1688/SPU），保留已填 SPU、分类、勾选、标记
    - 新出现的商品：追加
    - 不在本次订单里的旧记录：保留（不删除）
    """
    d = load_data()
    if not d["products"]:
        sync_from_excel()
        d = load_data()
    inq_by_url = {norm(p["url"]): p for p in d["products"]}
    inq_by_inq = {p["inq"]: p for p in d["products"] if p.get("inq")}
    # 标题近似匹配：订单里的名称与库里的名称高度相似也算命中（防链接不一致）
    libs = [p for p in d["products"] if p.get("url")]

    def lib_hit(k, p):
        h = inq_by_url.get(k)
        if h:
            return h
        nm = (p.get("names") or [""])[0].strip()
        if nm:
            import difflib
            best, bestr = None, 0.0
            for lp in libs:
                for ln in (lp.get("names") or []):
                    r = difflib.SequenceMatcher(None, nm.lower()[:120], (ln or "").lower()[:120]).ratio()
                    if r > bestr:
                        best, bestr = lp, r
            if bestr >= 0.86:
                return best
        return None

    prods = read_order(path)
    today = datetime.date.today().isoformat()

    def pkey(rec):
        return rec.get("inq") or norm(rec.get("url"))

    # 旧记录只用于继承已填内容，不进入最终列表
    pairs_map = {}
    for op in d.get("pairs", []):
        if op.get("inq"):
            pairs_map[op["inq"]] = op
        if op.get("url"):
            pairs_map.setdefault(norm(op["url"]), op)
    news_map = {}
    for op in d.get("news", []):
        if op.get("url"):
            news_map[norm(op["url"])] = op

    added_pairs = added_news = updated = not_inq = with_sku = 0
    pairs_out, news_out = [], []
    img_auto = []          # 靠图片/规格自动认出是库内同款的（跨站点不同语言标题）
    for k, p in prods.items():
        hit = lib_hit(k, p)
        how = "url/title" if hit else ""
        dist = None
        if not hit:
            # 通道1：图片极像
            if p.get("img"):
                try:
                    hit, dist = similar.find_same_by_image(p["img"], libs)
                except Exception:
                    hit, dist = None, None
                if hit is not None:
                    how = "image"
            # 通道2：标题规格参数一致（同款不同配色，图片差异大）
            if not hit:
                try:
                    hit2, info2 = similar.find_same_by_title_spec(
                        p.get("names") or [], libs)
                except Exception:
                    hit2, info2 = None, None
                if hit2 is not None:
                    hit, dist, how = hit2, info2, "spec"
        if hit is not None and how in ("image", "spec"):
            img_auto.append({
                "name": (p.get("names") or [""])[0][:70],
                "matched": (hit.get("names") or [""])[0][:70],
                "spu": hit.get("spu") or "",
                "how": how,
                "dist": dist,
            })
        if not hit:
            # ---- 未询盘 → 新品（只在本次订单内）----
            not_inq += 1
            prev = news_map.get(k) or {}
            if prev:
                updated += 1
            else:
                added_news += 1
            news_out.append({
                "url": p["url"], "name": (p["names"] or [""])[0],
                "names": p["names"] or prev.get("names") or [],
                "img": p["img"] or prev.get("img") or "",
                "skus": p["skus"] or prev.get("skus") or [],
                "l1": prev.get("l1", "") or "", "l2": prev.get("l2", "") or "",
                "l3": prev.get("l3", "") or "",
                "filled_spu": prev.get("filled_spu", "") or "",
                "mark": prev.get("mark"), "checked": prev.get("checked", True),
            })
            continue
        if p["skus"]:
            with_sku += 1
            continue
        # ---- 已询盘无 SKU → 待采集配对（只在本次订单内）----
        prev = pairs_map.get(hit["inq"]) or pairs_map.get(k) or {}
        merged_names = hit["names"] + [n for n in p["names"] if n not in hit["names"]]
        iq = hit["inq"] or ""
        spu_key = (hit.get("spu") or "").strip()
        # 同一产品只留一条：优先按 SPU，其次按询盘号，最后按规范化链接
        exist = None
        for x in pairs_out:
            if spu_key and (x.get("spu") or "").strip() == spu_key:
                exist = x
                break
            if iq and (x.get("inq") or "") == iq:
                exist = x
                break
            if p["url"] and norm(x.get("url")) == norm(p["url"]):
                exist = x
                break
        if exist is not None:
            for n in merged_names:
                if n not in exist["names"]:
                    exist["names"].append(n)
            if not exist.get("url") and p["url"]:
                exist["url"] = p["url"]
            if not exist.get("img") and (hit["img"] or p["img"]):
                exist["img"] = hit["img"] or p["img"]
            alt = exist.setdefault("alt_urls", [])
            if p["url"] and p["url"] != exist.get("url") and p["url"] not in alt:
                alt.append(p["url"])
            inqs = exist.setdefault("inqs", [])
            for y in ([exist.get("inq")] + [iq] if iq else []):
                if y and y not in inqs:
                    inqs.append(y)
            if prev and not (exist.get("filled_spu") or "").strip():
                exist["filled_spu"] = prev.get("filled_spu", "") or ""
            updated += 1
            continue
        if prev:
            updated += 1
        else:
            added_pairs += 1
        pairs_out.append({
            "url": p["url"] or prev.get("url") or "",
            "names": merged_names,
            "inq": iq, "spu": spu_key or prev.get("spu") or "",
            "inqs": [iq] if iq else [],
            "u1688": hit["u1688"] or prev.get("u1688") or [],
            "img": hit["img"] or p["img"] or prev.get("img") or "",
            "filled_spu": prev.get("filled_spu", "") or "",
            "mark": prev.get("mark"), "note": prev.get("note") or "",
        })

    # 已填过 SPU 的配对（人工处理过的成果）即使不在本次订单里也保留，避免白干
    in_order_urls = {norm(x.get("url")) for x in pairs_out if x.get("url")}
    in_order_inq = {x.get("inq") for x in pairs_out if x.get("inq")}
    for op in d.get("pairs", []):
        if not (op.get("filled_spu") or "").strip():
            continue
        u = norm(op.get("url"))
        if (u and u in in_order_urls) or (op.get("inq") and op.get("inq") in in_order_inq):
            continue
        pairs_out.append(op)

    pairs = pairs_out
    pairs.sort(key=lambda x: (0 if x.get("spu") else 1, x.get("inq") or ""))
    news = news_out
    d["pairs"] = pairs
    d["news"] = news
    d["pairs_date"] = today
    if sb.enabled():
        try:
            sb.save_pairs(today, pairs)
            sb.save_news(today, news)
            invalidate_cache()
        except Exception as e:
            print("[Supabase 写入失败] %s" % e)
    save_data(d)
    return {"ok": True, "total": len(prods), "inq": len(prods) - not_inq,
            "with_sku": with_sku, "pairs": len(pairs), "not_inq": not_inq,
            "news": len(news), "added_pairs": added_pairs,
            "added_news": added_news, "updated": updated,
            "img_auto": len(img_auto), "img_auto_detail": img_auto[:20]}


def import_quotes(path):
    """导入询盘导出，把新品追加入四列表（并备份）。"""
    ws = openpyxl.load_workbook(path, data_only=True).active
    qc = {str(ws.cell(row=1, column=i).value).strip(): i
          for i in range(1, ws.max_column + 1) if ws.cell(row=1, column=i).value}
    need = ["商品ID", "商户商品链接"]
    for n in need:
        if n not in qc:
            return {"ok": False, "msg": "询盘文件缺少列：%s" % n}

    def g(r, col):
        i = qc.get(col)
        if not i:
            return ""
        v = ws.cell(row=r, column=i).value
        return str(v).strip() if v not in (None, "") else ""

    # 读取所有询盘商品
    quotes = []
    for r in range(2, ws.max_row + 1):
        gid = g(r, "商品ID")
        if not gid:
            continue
        raw1688 = g(r, "1688链接")
        links = []
        for x in raw1688.split("\n"):
            x = x.strip().split("?")[0]
            if x.startswith("http") and x not in links:
                links.append(x)
        quotes.append({
            "inq": gid,
            "name": g(r, "商品名称"),
            "url": g(r, "商户商品链接"),
            "links": links,
            "img": g(r, "需求商品图链接") or g(r, "采集商品图链接"),
            "l1": g(r, "*一级分类"),
            "l2": g(r, "二级分类"),
            "l3": g(r, "三级分类"),
        })

    # 现有四列表
    wb = openpyxl.load_workbook(INQ_FILE)
    ws2 = wb.active
    ex_url, ex_iq = set(), set()
    last = 3
    for r in range(4, ws2.max_row + 1):
        if not ws2.cell(row=r, column=2).value:
            continue
        last = r
        ex_url.add(norm(ws2.cell(row=r, column=3).value))
        iq = str(ws2.cell(row=r, column=4).value or "").strip()
        if iq:
            ex_iq.add(iq)

    todo = [q for q in quotes if norm(q["url"]) not in ex_url and q["inq"] not in ex_iq]
    if not todo:
        pa, nr = quote_followup(quotes)
        return {"ok": True, "added": 0, "total": len(quotes), "msg": "没有新产品需要导入",
                "pairs_added": pa, "news_removed": nr}

    bak = backup_file(INQ_FILE)
    start = last + 1
    for i, q in enumerate(todo):
        r = start + i
        ws2.cell(row=r, column=2, value=q["name"])
        ws2.cell(row=r, column=3, value=q["url"])
        ws2.cell(row=r, column=4, value=q["inq"])
        ws2.cell(row=r, column=5, value=None)
        ws2.cell(row=r, column=6, value="\n".join(q["links"]) if q["links"] else None)
        for c in range(1, 7):
            cell = ws2.cell(row=r, column=c)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = Border(left=Side(style="thin"), right=Side(style="thin"),
                                 top=Side(style="thin"), bottom=Side(style="thin"))
        ws2.row_dimensions[r].height = 84
    total = last + len(todo)

    tmp = INQ_FILE + ".new.xlsx"
    wb.save(tmp)

    # 下载图片嵌入
    cache = os.path.join(BASE, "_imgcache")
    os.makedirs(cache, exist_ok=True)
    got = []
    for i, q in enumerate(todo):
        u = q["img"]
        if not u.startswith("http"):
            continue
        key = re.sub(r"[^A-Za-z0-9]", "_", u)[-90:] + ".img"
        fp = os.path.join(cache, key)
        if not (os.path.exists(fp) and os.path.getsize(fp) > 500):
            try:
                req = urllib.request.Request(u, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=20) as resp:
                    data = resp.read()
                if len(data) < 500:
                    continue
                open(fp, "wb").write(data)
            except Exception:
                continue
        got.append((start + i, fp))
    if got:
        wb2 = openpyxl.load_workbook(tmp)
        ws3 = wb2.active
        for r, fp in got:
            try:
                im = XLImage(fp)
                ratio = min(72.0 / im.width, 72.0 / im.height)
                im.width, im.height = int(im.width * ratio), int(im.height * ratio)
                ws3.add_image(im, "A%d" % r)
            except Exception:
                pass
        wb2.save(tmp)

    try:
        os.replace(tmp, INQ_FILE)
    except PermissionError:
        os.remove(tmp)
        return {"ok": False, "msg": "四列表被占用，请关闭后重试"}

    sync_from_excel()

    # 新品同步到 Supabase products 表（否则云端产品库刷新后看不到）
    if sb.enabled() and todo:
        try:
            items = []
            for q in todo:
                m = re.search(r"INQ-(\d+)", q["inq"] or "")
                items.append({
                    "row": None, "names": [q["name"]], "url": q["url"],
                    "inq": q["inq"], "batch": m.group(1)[-4:] if m else "",
                    "spu": "", "spus": [], "u1688": q["links"], "img": q["img"],
                })
            sb.upsert_products(items)
            invalidate_cache()
        except Exception as e:
            print("[Supabase 新品同步失败]", e)

    pa, nr = quote_followup(quotes)
    return {"ok": True, "added": len(todo), "total": len(quotes),
            "imgs": len(got), "backup": bak, "total_rows": total,
            "pairs_added": pa, "news_removed": nr}


def quote_followup(quotes):
    """本批已询盘产品：无 SPU 的补进「待采集配对」，并从「未询盘新品」剔除。
    返回 (新增配对数, 剔除新品数)。重复上传同一导出也安全（按询盘号/链接去重）。"""
    d = load_data()
    pairs = d.get("pairs", [])
    have_inq = {p.get("inq") for p in pairs}
    have_url = {norm(p.get("url")) for p in pairs}
    by_inq = {p["inq"]: p for p in d.get("products", []) if p.get("inq")}
    by_url = {norm(p["url"]): p for p in d.get("products", []) if p.get("url")}
    added_pairs = []
    for q in quotes:
        pr = by_inq.get(q["inq"]) or by_url.get(norm(q["url"]))
        if not pr:
            continue
        if (pr.get("spu") or "").strip():          # 已有 SPU，无需采集
            continue
        if q["inq"] in have_inq or norm(q["url"]) in have_url:
            continue
        added_pairs.append({
            "url": q["url"] or pr.get("url", ""),
            "names": pr.get("names") or [q["name"]],
            "inq": q["inq"], "spu": "",
            "u1688": pr.get("u1688") or q["links"],
            "img": pr.get("img") or q["img"],
            "filled_spu": "", "note": "",
        })
    if added_pairs:
        d["pairs"] = pairs + added_pairs
        d["pairs_date"] = datetime.date.today().isoformat()

    q_urls = {norm(q["url"]) for q in quotes if q["url"]}
    news = d.get("news", [])
    keep = [n for n in news if norm(n.get("url")) not in q_urls]
    removed_news = len(news) - len(keep)
    if removed_news:
        if sb.enabled():
            for n in news:
                if norm(n.get("url")) in q_urls and n.get("id"):
                    try:
                        sb.delete("news", {"id": n["id"]})
                        sb.log_op("import", "news_quoted_removed",
                                  {"name": (n.get("names") or [n.get("name", "")])[0][:80]})
                    except Exception as e:
                        print("[新品剔除失败]", e)
        d["news"] = keep

    if added_pairs or removed_news:
        save_data(d)
        if sb.enabled() and added_pairs:
            try:
                sb.save_pairs(d["pairs_date"], d["pairs"])
            except Exception as e:
                print("[Supabase 配对写入失败]", e)
        invalidate_cache()
    return len(added_pairs), removed_news


def export_demand():
    """导出询盘需求表（勾选的未询盘产品）。"""
    d = load_data()
    news = [x for x in d.get("news", []) if x.get("checked", True)]
    if not news:
        return {"ok": False, "msg": "没有勾选的产品"}
    tpl = None
    for f in os.listdir(BASE):
        if f.startswith("template_center_inquiry_demand-entry_") and f.endswith(".xlsx"):
            tpl = os.path.join(BASE, f)
            break
    if not tpl:
        return {"ok": False, "msg": "找不到需求表模板"}

    wb = openpyxl.load_workbook(tpl)
    ws = wb["商品导入"]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    for i, x in enumerate(news, start=2):
        nm = x["name"][:100].rstrip(" –-,")
        ws.cell(row=i, column=1).value = "普货"
        ws.cell(row=i, column=2).value = nm
        ws.cell(row=i, column=3).value = x.get("img", "")
        ws.cell(row=i, column=4).value = 1
        ws.cell(row=i, column=5).value = x.get("l1") or None
        ws.cell(row=i, column=6).value = x.get("l2") or None
        ws.cell(row=i, column=7).value = x.get("l3") or None
        ws.cell(row=i, column=9).value = x["url"]
        ws.cell(row=i, column=13).value = "有单"
        ws.cell(row=i, column=2).alignment = Alignment(vertical="center", wrap_text=True)
        ws.cell(row=i, column=3).alignment = Alignment(vertical="center", wrap_text=True)
    for c, w in {"A": 10, "B": 42, "C": 58, "D": 12, "E": 12,
                 "F": 16, "G": 22, "I": 76, "M": 10}.items():
        ws.column_dimensions[c].width = w

    out = os.path.join(BASE, "Mico杂货铺询盘需求表_%s.xlsx"
                       % datetime.datetime.now().strftime("%Y%m%d_%H%M"))
    try:
        wb.save(out)
    except Exception as e:
        return {"ok": False, "msg": str(e)}
    return {"ok": True, "file": os.path.basename(out), "count": len(news)}


def export_pairs():
    d = load_data()
    pairs = d.get("pairs", [])
    if not pairs:
        return {"ok": False, "msg": "没有配对数据"}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "待采集配对产品"
    for i, h in enumerate(["产品图", "产品名称（客户标题）", "客户链接", "询盘号", "SPU", "1688链接"], 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=Side(style="thin"), right=Side(style="thin"),
                          top=Side(style="thin"), bottom=Side(style="thin"))
    ws.row_dimensions[1].height = 26
    for cc, w in {"A": 15, "B": 58, "C": 62, "D": 25, "E": 16, "F": 40}.items():
        ws.column_dimensions[cc].width = w
    for i, p in enumerate(pairs):
        r = i + 2
        ws.row_dimensions[r].height = max(84, 16 * len(p["names"]) + 26)
        ws.cell(row=r, column=2, value="\n".join(p["names"]))
        ws.cell(row=r, column=3, value=p["url"])
        ws.cell(row=r, column=4, value=p["inq"] or None)
        ws.cell(row=r, column=5, value=p.get("filled_spu") or p.get("spu") or None)
        _u = sb.uniq_urls(p.get("u1688"))
        ws.cell(row=r, column=6, value="\n".join(_u) if _u else None)
        for c in range(1, 7):
            cell = ws.cell(row=r, column=c)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = Border(left=Side(style="thin"), right=Side(style="thin"),
                                 top=Side(style="thin"), bottom=Side(style="thin"))
    date = d.get("pairs_date", datetime.date.today().isoformat())
    out = os.path.join(BASE, "%s-待采集配对产品.xlsx" % date)
    tmp = os.path.join(BASE, "_tmp_%s.xlsx" % datetime.datetime.now().strftime("%H%M%S"))
    wb.save(tmp)

    cache = os.path.join(BASE, "_imgcache")
    os.makedirs(cache, exist_ok=True)
    got = []
    for i, p in enumerate(pairs):
        u = p.get("img") or ""
        if not u.startswith("http"):
            continue
        key = re.sub(r"[^A-Za-z0-9]", "_", u)[-90:] + ".img"
        fp = os.path.join(cache, key)
        if not (os.path.exists(fp) and os.path.getsize(fp) > 500):
            try:
                req = urllib.request.Request(u, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=20) as resp:
                    data = resp.read()
                if len(data) < 500:
                    continue
                open(fp, "wb").write(data)
            except Exception:
                continue
        got.append((i + 2, fp))
    if got:
        wb2 = openpyxl.load_workbook(tmp)
        ws2 = wb2.active
        for r, fp in got:
            try:
                im = XLImage(fp)
                ratio = min(72.0 / im.width, 72.0 / im.height)
                im.width, im.height = int(im.width * ratio), int(im.height * ratio)
                ws2.add_image(im, "A%d" % r)
            except Exception:
                pass
        wb2.save(tmp)
    try:
        os.replace(tmp, out)
    except PermissionError:
        os.remove(tmp)
        return {"ok": False, "msg": "文件被占用，请关闭 %s" % os.path.basename(out)}
    return {"ok": True, "file": os.path.basename(out), "count": len(pairs), "imgs": len(got)}


def save_spu_supabase(arr, editor):
    """把填写的 SPU 写入 Supabase pairs 表（按 id / 询盘号 / 链接三重匹配）。
    返回 (写入数, 失败数, 错误列表)。多人协作必须确认真的写进去了，不能静默失败。"""
    if not sb.enabled():
        return 0, 0, ["未配置 SUPABASE_URL / SUPABASE_ANON_KEY"]
    rows = sb.get_pairs()
    by_id, by_inq, by_url = {}, {}, {}
    for r in rows:
        if r.get("id") is not None:
            by_id[r["id"]] = r
        if r.get("inq"):
            by_inq[r["inq"]] = r
        if r.get("url"):
            by_url[norm(r["url"])] = r
    lst = load_data().get("pairs", [])
    ok_n = fail_n = 0
    errs = []
    for x in arr:
        i2 = x.get("idx")
        spu = (x.get("spu") or "").strip()
        if not spu or not isinstance(i2, int) or not (0 <= i2 < len(lst)):
            continue
        p = lst[i2]
        row = None
        if p.get("id") is not None:
            row = by_id.get(p["id"])
        if row is None and p.get("inq"):
            row = by_inq.get(p["inq"])
        if row is None and p.get("url"):
            row = by_url.get(norm(p["url"]))
        if row is None:
            fail_n += 1
            errs.append("找不到云端记录：%s" % ((p.get("names") or [""])[0][:40]))
            continue
        try:
            sb.update_pair_spu(row["id"], spu, editor)
            p["id"] = row["id"]
            ok_n += 1
            # 同一 SPU 的多个客户外语名全部记录，方便以后按 SPU 查询
            names = p.get("names") or []
            sb.put_spu_names([{"spu": s, "name": n, "url": p.get("url") or "",
                               "inq": p.get("inq") or ""}
                              for s in [y.strip() for y in re.split(r"[,\s/;]+", spu) if y.strip()]
                              for n in names])
        except Exception as e:
            fail_n += 1
            errs.append("写入失败(%s)：%s" % (row["id"], str(e)[:80]))
    invalidate_cache()
    return ok_n, fail_n, errs


def record_edit(pairs, editor, arr):
    """记录谁在什么时候改了什么（edit_log 并行写，省网络往返时间）。"""
    if sb.enabled():
        try:
            jobs = []
            for x in arr:
                i2 = x.get("idx")
                spu = (x.get("spu") or "").strip()
                if not spu or not isinstance(i2, int) or not (0 <= i2 < len(pairs)):
                    continue
                old = pairs[i2].get("spu") or ""
                nm = (pairs[i2].get("names") or [""])[0][:80]
                jobs.append((pairs[i2].get("id"), old, spu, nm))
            if jobs:
                from concurrent.futures import ThreadPoolExecutor
                with ThreadPoolExecutor(max_workers=min(6, len(jobs))) as ex:
                    list(ex.map(
                        lambda j: sb.log_edit(editor, "pairs", j[0], j[1], j[2], j[3]),
                        jobs))
        except Exception as e:
            print("[Supabase 日志写入失败] %s" % e)
    d = load_data()
    log = d.setdefault("edit_log", [])
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for x in arr:
        i = x.get("idx")
        spu = (x.get("spu") or "").strip()
        if not spu or not isinstance(i, int) or not (0 <= i < len(pairs)):
            continue
        nm = (pairs[i].get("names") or [""])[0][:60]
        log.append({"time": ts, "editor": editor, "idx": i,
                    "spu": spu, "name": nm})
    d["edit_log"] = log[-500:]      # 只保留最近 500 条
    save_data(d)


def apply_spu_to_master(user_name="", sync_pairs=True):
    """sync_pairs=False 时跳过 pairs 重复写入（save-spu 路径已由 save_spu_supabase 写过）。"""
    d = load_data()
    filled = [p for p in d.get("pairs", []) if (p.get("filled_spu") or "").strip()]
    if not filled:
        return {"ok": False, "msg": "配对表里没有填写任何 SPU"}
    bak = backup_file(INQ_FILE)
    wb = openpyxl.load_workbook(INQ_FILE)
    ws = wb.active
    by_url, by_iq = {}, {}
    for r in range(4, ws.max_row + 1):
        if not ws.cell(row=r, column=2).value:
            continue
        by_url[norm(ws.cell(row=r, column=3).value)] = r
        iq = str(ws.cell(row=r, column=4).value or "").strip()
        if iq:
            by_iq[iq] = r
    upd = skip = 0
    for p in filled:
        r = by_iq.get(p.get("inq")) or by_url.get(norm(p.get("url")))
        if not r:
            skip += 1
            continue
        ws.cell(row=r, column=5).value = p["filled_spu"].strip()
        ws.cell(row=r, column=5).font = Font(name="Calibri", size=10)
        upd += 1
    wb.save(INQ_FILE)
    if sb.enabled():
        try:
            # 配对填的 SPU 写回产品库：有就更新、没有就新增（这样填过的产品永久留在库内，
            # 后续导入订单按产品库匹配就能认出它是「已询盘」）
            plist = sb.get_products()
            by_inq, by_url = {}, {}
            for pr in plist:
                if pr.get("inq"):
                    by_inq[pr["inq"]] = pr
                if pr.get("url"):
                    by_url[sb.norm_url(pr["url"])] = pr
            new_rows = []
            for p in filled:
                spu_txt = p["filled_spu"].strip()
                ex = None
                if p.get("inq"):
                    ex = by_inq.get(p["inq"])
                if ex is None and p.get("url"):
                    ex = by_url.get(sb.norm_url(p["url"]))
                if ex is not None:
                    try:
                        sb.update("products", {"id": ex["id"]},
                                  {"spu": spu_txt, "spus": sb.split_spu(spu_txt)})
                    except Exception:
                        pass
                else:
                    m = re.search(r"INQ-(\d+)", p.get("inq") or "")
                    new_rows.append({
                        "row": None,
                        "names": p.get("names") or [""],
                        "url": p.get("url") or "",
                        "inq": p.get("inq") or None,
                        "batch": m.group(1)[-4:] if m else None,
                        "spu": spu_txt,
                        "spus": sb.split_spu(spu_txt),
                        "u1688": p.get("u1688") or [],
                        "img": p.get("img") or None,
                    })
            if new_rows:
                sb.upsert_products(new_rows)
                print("[产品库新增] 配对转入 %d 条" % len(new_rows))
            # 同步 pairs 表里的人工填写值（save-spu 路径已写过，跳过省一次往返）
            if sync_pairs:
                for p in filled:
                    if p.get("id"):
                        try:
                            sb.update_pair_spu(p["id"], p["filled_spu"].strip(), user_name)
                        except Exception:
                            pass
        except Exception as e:
            print("[Supabase SPU 同步失败] %s" % e)
    invalidate_cache()
    if not sb.enabled():
        sync_from_excel()
    return {"ok": True, "updated": upd, "skipped": skip, "backup": bak}


# ================= 认证 =================
import base64
import hashlib

COOKIE_NAME = "mico_auth"


def make_token(user):
    raw = "%s|%s" % (user, APP_PASSWORD)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:32]


def parse_cookie(header):
    out = {}
    if not header:
        return out
    for part in header.split(";"):
        part = part.strip()
        if "=" in part:
            k, v = part.split("=", 1)
            out[k.strip()] = v.strip()
    return out


# ================= 渲染 =================
def esc(s):
    return html.escape(str(s))


# ---------- 三级分类（平台字典） ----------
_CAT_TREE = None


def cat_tree():
    global _CAT_TREE
    if _CAT_TREE is None:
        fp = os.path.join(BASE, "cat_tree.json")
        try:
            with open(fp, encoding="utf-8") as f:
                _CAT_TREE = json.load(f)
        except Exception as e:
            print("[分类字典加载失败] %s" % e)
            _CAT_TREE = {}
    return _CAT_TREE


def _opts(values, cur):
    out = []
    for v in values:
        out.append('<option value="%s"%s>%s</option>'
                   % (esc(v), ' selected' if v == cur else '', esc(v)))
    return "".join(out)


def cat_opts_l1(cur=""):
    return _opts(sorted(cat_tree().keys()), cur)


def cat_opts_l2(l1, cur=""):
    if not l1:
        return ""
    return _opts(sorted((cat_tree().get(l1) or {}).keys()), cur)


def cat_opts_l3(l1, l2, cur=""):
    if not l1 or not l2:
        return ""
    return _opts(sorted(((cat_tree().get(l1) or {}).get(l2) or [])), cur)


def cn_badge(cn):
    """中文品名徽章：点击后切到产品库并按该品名筛选。"""
    if not cn:
        return ""
    return ('<b class="cnbadge" title="点击查看全部「%s」同款" '
            'onclick="byCn(\'%s\')">%s</b>' % (esc(cn), esc(cn), esc(cn)))


def product_card(p, i=0):
    name = p["names"][0] if p["names"] else ""
    cn = similar.cn_for(p)
    cnb = cn_badge(cn)
    extra = len(p["names"]) - 1
    if p.get("multi_spu"):
        spu_html = "".join('<b class="spu multi">%s</b>' % esc(s) for s in p["spus"])
        spu_html += '<b class="tag">多</b>'
    elif p["spu"]:
        spu_html = '<b class="spu">%s</b>' % esc(p["spu"])
    else:
        spu_html = '<b class="spu none">无 SPU</b>'
    img = ('<img src="%s" loading="lazy" decoding="async">' % esc(p["img"])) if p["img"] else '<span class="nopic">无图</span>'
    urls = [u for u in (p.get("urls") or ([p["url"]] if p.get("url") else [])) if u and u.startswith("http")]
    inqs = [x for x in (p.get("inqs") or ([p["inq"]] if p.get("inq") else [])) if x]
    if len(urls) > 1:
        site_rows = "".join(
            '<div class="l1688"><a href="%s" target="_blank" rel="noopener" class="siteurl">&#127761; %s</a></div>'
            % (esc(u), esc(re.sub(r"^https?://(www\.)?", "", u)[:38])) for u in urls)
        link = '<div class="links sitelist">%s</div>' % site_rows
    else:
        link = ('<a href="%s" target="_blank" rel="noopener">站点</a>' % esc(urls[0])) if urls else ""
    l1688 = ('<a href="%s" target="_blank" rel="noopener">1688</a>' % esc((p.get("u1688") or [""])[0])) if (p.get("u1688") or []) else ""
    if len(p["names"]) > 1:
        names_html = "".join('<div class="nmrow">%s</div>' % esc(n[:100]) for n in p["names"])
        title_html = '<div class="pn names">%s %s</div>' % (names_html, cnb)
    else:
        title_html = '<div class="pn" title="%s">%s%s %s</div>' % (
            esc(" ／ ".join(p["names"])), esc(name[:90]),
            ('<i>+%d</i>' % extra) if extra > 0 else "", cnb)
    return (
        '<div class="pc" data-n="%s" data-s="%s" data-i="%s" data-m="%s" data-p="%s" data-q="%s" data-cn="%s" data-row="%s">'
        '<div class="pi">%s</div><div class="pb">'
        '%s'
        '<div class="ps">%s</div>'
        '<div class="pf"><span class="iq">%s</span></div>'
        '<div class="spuedit"><input class="libspu" data-inq="%s" data-row="%s" value="%s" placeholder="多个 SPU 用逗号分隔">'
        '<button class="libsave" onclick="saveLibSpu(this)">存</button>'
        '<button class="libadd" onclick="addLibSpu(this)" title="在已有 SPU 后追加">+</button></div>'
        '<div class="pl">%s %s'
        '<button class="simbtn" onclick="sim(%d,\'lib\')">相似款</button></div></div></div>' % (
            esc(name.lower() + " " + cn), esc(p["spu"].lower()), esc((inqs[0] if inqs else "").lower()),
            "1" if p.get("multi_spu") else "0",
            "0" if p["spu"] else "1",
            "1" if not inqs else "0",
            esc(cn), esc(str(p.get("row") or "")),
            img, title_html,
            spu_html, esc("、".join(x[-6:] for x in inqs) if inqs else "—"),
            esc(inqs[0] if inqs else ""), esc(str(p.get("row") or "")), esc(p["spu"]),
            link, l1688, i))


def pair_card(p, i=0):
    """p 为 merge_by_spu 的合并组（含 _members/_idx）。"""
    img = ('<img src="%s" loading="lazy">' % esc(p["img"])) if p.get("img") else '<span class="nopic">无图</span>'
    cur = p.get("filled_spu") or p.get("spu") or ""
    links = sb.uniq_urls(p.get("u1688"))
    rows = ""
    if links:
        rows += "".join(
            '<div class="l1688"><input class="i1688" readonly value="%s" '
            'onclick="this.select();document.execCommand(\'copy\')" '
            'onfocus="this.select()" title="点击即复制">'
            '<a href="%s" target="_blank" rel="noopener">打开</a></div>'
            % (esc(u), esc(u)) for u in links)
    else:
        rows += '<div class="l1688"><span class="nolink">无 1688 链接</span></div>'
    urls = [u for u in (p.get("urls") or [p.get("url") or ""]) if u and u.startswith("http")]
    for u in urls:
        short = re.sub(r"^https?://(www\.)?", "", u)
        if len(short) > 34:
            short = short[:34] + "…"
        rows += ('<div class="l1688 lsite"><a href="%s" target="_blank" rel="noopener" '
                 'title="%s" class="siteurl">&#127761; %s</a></div>'
                 % (esc(u), esc(u), esc(short)))
    l1688 = '<div class="links">%s</div>' % rows
    names = [n for n in (p.get("names") or []) if n]
    if len(names) > 1:
        names_html = "".join('<div class="nmrow">%s</div>' % esc(n[:100]) for n in names)
        title_html = '<div class="pn names">%s</div>' % names_html
    else:
        title_html = '<div class="pn">%s</div>' % esc((names[0] if names else "")[:84])
    cn = similar.cn_for(p)
    cnb = cn_badge(cn)
    inqs = [x for x in (p.get("inqs") or []) if x]
    inq_txt = "、".join(x[-6:] for x in inqs) if inqs else "—"
    idxs = p.get("_idx") or [i]
    idx_attr = ",".join(str(x) for x in idxs)
    return (
        '<div class="paircard" data-inq="%s" data-url="%s" data-cn="%s" data-idxs="%s"><div class="pi">%s</div><div class="pb">'
        '%s %s'
        '<div class="ps"><span class="iq">%s</span>%s</div>'
        '%s'
        '<div class="inputrow"><label>SPU</label>'
        '<input class="spin" data-idx="%d" data-idxs="%s" value="%s" placeholder="多个用逗号分隔">'
        '<button class="spuq" onclick="spuNames(%d)" title="查看该 SPU 已记录的所有客户产品名">查名</button>'
        '<button class="libadd" onclick="addPairSpu(%d)" title="追加一个 SPU，不覆盖已有的">+</button></div>'
        '<div style="display:flex;gap:6px;margin-top:4px">'
        '<button class="simbtn%s" onclick="sim(%d,\'pair\')">相似款比对</button>'
        '<button class="simbtn savebtn" onclick="saveOne(%d)">保存</button>'
        '</div>'
        '</div></div>' % (
            esc((inqs[0] if inqs else "")), esc(urls[0] if urls else ""), esc(cn), esc(idx_attr),
            img, title_html, cnb,
            esc(inq_txt),
            ('<i class="tag">合并%d</i>' % len(idxs)) if len(idxs) > 1 else "",
            l1688, i, esc(idx_attr), esc(cur), i, i,
            " done" if p.get("mark") == "same" else "", i, i))


def news_card(p, i=0):
    img = ('<img src="%s" loading="lazy">' % esc(p["img"])) if p.get("img") else '<span class="nopic">无图</span>'
    checked = "checked" if p.get("checked", True) else ""
    cn = similar.cn_for(p)
    cnb = cn_badge(cn)
    spu_v = p.get("filled_spu") or ""
    names = [n for n in (p.get("names") or [p.get("name") or ""]) if n]
    if len(names) > 1:
        names_html = "".join('<div class="nmrow">%s</div>' % esc(n[:100]) for n in names)
        title_html = '<div class="pn names">%s %s</div>' % (names_html, cnb)
    else:
        nm0 = names[0] if names else ""
        title_html = '<div class="pn" title="%s">%s %s</div>' % (esc(nm0), esc(nm0[:80]), cnb)
    urls = [u for u in (p.get("urls") or [p.get("url") or ""]) if u and u.startswith("http")]
    if len(urls) > 1:
        site_html = "".join(
            '<div class="l1688"><a href="%s" target="_blank" rel="noopener" class="siteurl">&#127761; %s</a></div>'
            % (esc(u), esc(re.sub(r"^https?://(www\.)?", "", u)[:38])) for u in urls)
        site_html = '<div class="links sitelist">%s</div>' % site_html
    else:
        site_html = '<a href="%s" target="_blank" rel="noopener">站点</a>' % (esc(urls[0]) if urls else "")
    idxs = p.get("_idx") or [i]
    idx_attr = ",".join(str(x) for x in idxs)
    return (
        '<div class="newscard" data-idx="%d" data-idxs="%s">'
        '<label class="chk"><input type="checkbox" data-idx="%d" %s></label>'
        '<div class="pi">%s</div><div class="pb">'
        '%s'
        '<div class="cat3">'
        '<select class="ci" data-idx="%d" data-k="l1" onchange="catL1(this)">'
        '<option value="">一级分类</option>%s</select>'
        '<select class="ci" data-idx="%d" data-k="l2" onchange="catL2(this)">'
        '<option value="">二级分类</option>%s</select>'
        '<select class="ci" data-idx="%d" data-k="l3">'
        '<option value="">三级分类</option>%s</select>'
        '</div>'
        '<div class="inputrow"><label>SPU</label>'
        '<input class="spn%s" data-idx="%d" data-idxs="%s" value="%s" placeholder="多个用逗号分隔">'
        '<button class="libadd" onclick="addNewsSpu(%d)" title="追加一个 SPU，不覆盖已有的">+</button></div>'
        '<div class="pl">%s'
        '<button class="simbtn%s" onclick="sim(%d,\'news\')">相似款</button>'
        '<button class="simbtn savebtn" onclick="saveNewsSpu(%d)">保存SPU</button></div>'
        '</div></div>' % (
            i, esc(idx_attr), i, checked, img, title_html,
            i, cat_opts_l1(p.get("l1", "")),
            i, cat_opts_l2(p.get("l1", ""), p.get("l2", "")),
            i, cat_opts_l3(p.get("l1", ""), p.get("l2", ""), p.get("l3", "")),
            " filled" if spu_v else "", i, esc(idx_attr), esc(spu_v), i,
            site_html,
            " done" if p.get("mark") == "same" else "", i, i))


LOGIN_TPL = """<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Mico 工作台 · 登录</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,"Segoe UI","Microsoft YaHei",sans-serif;
 background:linear-gradient(135deg,#1e293b,#334155 60%,#475569);
 min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.box{background:#fff;border-radius:16px;padding:38px 34px;width:100%;max-width:360px;
 box-shadow:0 20px 50px rgba(0,0,0,.28)}
h1{font-size:20px;font-weight:700;margin-bottom:6px;color:#1e293b}
.sub{font-size:13px;color:#94a3b8;margin-bottom:26px}
label{display:block;font-size:12.5px;color:#64748b;font-weight:600;margin:14px 0 6px}
input,select{width:100%;padding:11px 14px;border:1px solid #e2e8f0;border-radius:9px;
 font-size:14px;outline:none;font-family:inherit}
input:focus,select:focus{border-color:#4f46e5;box-shadow:0 0 0 3px rgba(79,70,229,.1)}
button{width:100%;margin-top:22px;padding:12px;border:none;border-radius:9px;
 background:#4f46e5;color:#fff;font-size:14.5px;font-weight:600;cursor:pointer;font-family:inherit}
button:hover{background:#4338ca}
.err{background:#fef2f2;color:#b91c1c;font-size:12.5px;padding:9px 12px;
 border-radius:8px;margin-bottom:14px;border:1px solid #fecaca}
</style></head><body>
<div class="box">
  <h1>Mico 选品工作台</h1>
  <span class="user">{{USER}} <a href="/logout" class="logout">退出</a></span>
  <div class="sub">请输入姓名和密码</div>
  {{ERR}}
  <form method="POST" action="/login">
    <label>我是谁</label>
    <select name="user">{{USERS}}</select>
    <label>密码</label>
    <input type="password" name="pwd" placeholder="输入团队密码" autofocus>
    <button type="submit">进入工作台</button>
  </form>
</div>
</body></html>"""


def page_login(err=""):
    opts = "".join('<option value="%s">%s</option>' % (esc(u), esc(u)) for u in USERS)
    e = '<div class="err">%s</div>' % esc(err) if err else ""
    return LOGIN_TPL.replace("{{USERS}}", opts).replace("{{ERR}}", e)


CURRENT_USER = {"name": ""}


def merge_by_spu(prods, with_idx=False):
    """相同 SPU 的卡片合并为一条展示记录（一个货号多站点多名称）。
    with_idx=True 时，记录组内成员在原始列表中的下标（保存时整组写入）。"""
    groups, order = {}, []
    for i, p in enumerate(prods):
        spus = [s for s in (p.get("spus") or ([p["spu"]] if p.get("spu") else [])) if s]
        if not spus:
            fs = (p.get("filled_spu") or "").strip()
            if fs:
                spus = sb.split_spu(fs)
        spus = [s for s in spus if s]
        key = spus[0] if len(spus) == 1 else None
        k = ("SPU", key) if key else ("ROW", id(p))
        if k not in groups:
            g = dict(p)
            g["names"] = list(p.get("names") or [])
            g["urls"] = [p.get("url") or ""]
            g["inqs"] = [p.get("inq") or ""]
            g["u1688"] = list(p.get("u1688") or [])
            g["_members"] = [p]
            g["_idx"] = [i]
            g["_key"] = key or ""
            groups[k] = g
            order.append(k)
            continue
        g = groups[k]
        for n in (p.get("names") or []):
            if n and n not in g["names"]:
                g["names"].append(n)
        u = p.get("url") or ""
        if u and u not in g["urls"]:
            g["urls"].append(u)
        iq = p.get("inq") or ""
        if iq and iq not in g["inqs"]:
            g["inqs"].append(iq)
        for x in (p.get("u1688") or []):
            if x and x not in g["u1688"]:
                g["u1688"].append(x)
        if not g.get("img") and p.get("img"):
            g["img"] = p["img"]
        if not (g.get("filled_spu") or "").strip() and (p.get("filled_spu") or "").strip():
            g["filled_spu"] = p["filled_spu"]
        g["_members"].append(p)
        g["_idx"].append(i)
    return [groups[k] for k in order]


def page_index():
    d = load_data()
    prods = d.get("products", [])
    pairs = d.get("pairs", [])
    n_spu = sum(1 for p in prods if p["spu"])
    n_multi = sum(1 for p in prods if p.get("multi_spu"))
    n_noiq = sum(1 for p in prods if not p["inq"])

    prod_groups = merge_by_spu(prods)
    pair_groups = merge_by_spu(pairs, with_idx=True)
    news = d.get("news", [])
    news_groups = merge_by_spu(news, with_idx=True)

    cards = "".join(product_card(p, i) for i, p in enumerate(prod_groups))
    pcards = "".join(pair_card(g, i) for i, g in enumerate(pair_groups))
    no_pairs = '<div class="empty">还没有配对数据<br><span>请在 ① 上传今日订单文件</span></div>' if not pairs else ""

    ncards = "".join(news_card(g, i) for i, g in enumerate(news_groups))
    no_news = '<div class="empty">还没有未询盘产品<br><span>请在 ① 上传今日订单文件</span></div>' if not news else ""

    return PAGE_TPL \
        .replace("{{USER}}", esc(CURRENT_USER.get("name", ""))) \
        .replace("{{TOTAL}}", str(len(prod_groups))) \
        .replace("{{NSPU}}", str(n_spu)) \
        .replace("{{NMULTI}}", str(n_multi)) \
        .replace("{{NNOIQ}}", str(n_noiq)) \
        .replace("{{NPAIRS}}", str(len(pair_groups))) \
        .replace("{{PER}}", str(PER_PAGE)) \
        .replace("{{CARDS}}", cards) \
        .replace("{{PCARDS}}", pcards) \
        .replace("{{EMPTY}}", no_pairs) \
        .replace("{{NCARDS}}", ncards) \
        .replace("{{NONEWS}}", no_news) \
        .replace("{{NNEWS}}", str(len(news_groups)))


PAGE_TPL = r"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Mico 选品工作台</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,"Segoe UI","Microsoft YaHei",sans-serif;
 background:#f0f2f5;color:#1f2937;font-size:14px;line-height:1.6}
header{background:linear-gradient(135deg,#1e293b,#334155 60%,#475569);
 color:#fff;padding:14px 22px;box-shadow:0 2px 10px rgba(0,0,0,.14)}
.hrow{max-width:1560px;margin:0 auto;display:flex;align-items:center;gap:20px;flex-wrap:wrap}
header h1{font-size:17px;font-weight:700;white-space:nowrap}
.user{font-size:12.5px;opacity:.8;white-space:nowrap}
.logout{color:#a5b4fc;text-decoration:none;margin-left:6px}
.logout:hover{text-decoration:underline}
.stats{display:flex;gap:18px;flex-wrap:wrap;margin-left:auto}
.st{text-align:center;min-width:58px}
.st b{display:block;font-size:18px;font-weight:700;line-height:1.2}
.st span{font-size:10.5px;opacity:.62}
main{max-width:1560px;margin:0 auto;padding:16px 22px 60px}
.tabs{display:flex;gap:7px;margin-bottom:14px;flex-wrap:wrap}
.tab{padding:7px 17px;border-radius:8px;background:#fff;border:1px solid #e2e8f0;
 cursor:pointer;font-size:13px;color:#64748b;font-weight:500}
.tab.on{background:#4f46e5;color:#fff;border-color:#4f46e5;font-weight:600}
.panel{display:none}.panel.on{display:block}
.bar{display:flex;gap:9px;margin-bottom:13px;flex-wrap:wrap;align-items:center}
.search{flex:1;min-width:200px;padding:9px 14px;border:1px solid #e2e8f0;
 border-radius:8px;font-size:13px;outline:none;font-family:inherit;background:#fff}
.search:focus{border-color:#a5b4fc}
.chip{padding:6px 14px;border-radius:18px;background:#fff;border:1px solid #e2e8f0;
 cursor:pointer;font-size:12px;color:#64748b;white-space:nowrap}
.chip.on{background:#eef2ff;border-color:#a5b4fc;color:#4338ca;font-weight:600}
.btn{padding:9px 19px;border-radius:8px;border:none;background:#4f46e5;color:#fff;
 font-size:13px;font-weight:600;cursor:pointer;font-family:inherit}
.btn:hover{background:#4338ca}
.btn.ghost{background:#fff;color:#4338ca;border:1px solid #c7d2fe}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:12px}
.pc,.paircard{background:#fff;border-radius:10px;overflow:hidden;
 box-shadow:0 1px 2px rgba(16,24,40,.07);display:flex;flex-direction:column}
.pc:hover{box-shadow:0 3px 10px rgba(16,24,40,.12)}
.pi{height:132px;background:#f8fafc;display:flex;align-items:center;justify-content:center;overflow:hidden}
.pi img{width:100%;height:100%;object-fit:contain;background:#fff}
.nopic{color:#cbd5e1;font-size:11px}
.pb{padding:9px 10px 10px;display:flex;flex-direction:column;gap:5px;flex:1}
.pn{font-size:12px;font-weight:600;color:#1e293b;line-height:1.38;
 display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;min-height:33px}
.pn i{color:#94a3b8;font-style:normal;font-size:10.5px}
.ps{display:flex;flex-wrap:wrap;gap:4px}
.spu{display:inline-block;font-family:ui-monospace,Consolas,monospace;font-size:10.5px;
 font-weight:600;padding:1px 6px;border-radius:4px;background:#ecfdf5;color:#047857;border:1px solid #a7f3d0}
.spu.multi{background:#fef3c7;color:#b45309;border-color:#fcd34d}
.spu.none{background:#fef2f2;color:#b91c1c;border-color:#fecaca}
.tag{font-size:9.5px;padding:1px 5px;border-radius:4px;background:#8b5cf6;color:#fff;font-weight:600}
.pf{margin-top:auto}
.iq{font-family:ui-monospace,Consolas,monospace;font-size:10.5px;color:#64748b}
.pl{display:flex;gap:8px;font-size:11px}
.pl a{color:#4f46e5;text-decoration:none;font-weight:500}
.pl a:hover{text-decoration:underline}
.inputrow{display:flex;align-items:center;gap:7px;margin-top:4px}
.inputrow label{font-size:11px;color:#64748b;font-weight:600}
.spin,.spn{flex:1;padding:6px 9px;border:1px solid #e2e8f0;border-radius:6px;
 font-family:ui-monospace,Consolas,monospace;font-size:11.5px;outline:none;min-width:0}
.spin:focus{border-color:#4f46e5}
.spn:focus{border-color:#4f46e5}
.spin.filled{border-color:#10b981;background:#f0fdf4}
.spn.filled{border-color:#10b981;background:#f0fdf4}
.drop{border:2px dashed #cbd5e1;border-radius:12px;padding:34px 22px;text-align:center;
 background:#fff;cursor:pointer}
.drop:hover,.drop.over{border-color:#4f46e5;background:#f5f3ff}
.drop .ico{font-size:32px;margin-bottom:8px}
.drop h3{font-size:14.5px;font-weight:600;margin-bottom:5px}
.drop p{font-size:12px;color:#94a3b8}
.summary{background:#fff;border-radius:11px;padding:16px 20px;margin-top:14px;
 box-shadow:0 1px 2px rgba(16,24,40,.07);display:none}
.summary.show{display:block}
.summary .srow{display:flex;gap:22px;flex-wrap:wrap;margin-bottom:11px}
.summary .si{font-size:12.5px;color:#475569}
.summary .si b{font-size:18px;color:#1e293b;margin-right:4px}
.empty{text-align:center;padding:50px 20px;color:#94a3b8;font-size:13.5px}
.empty span{font-size:12px;color:#cbd5e1}
.toast{position:fixed;bottom:24px;left:50%;transform:translateX(-50%) translateY(70px);
 background:#1e293b;color:#fff;padding:11px 22px;border-radius:9px;font-size:13px;
 opacity:0;transition:.25s;z-index:99}
.toast.show{opacity:1;transform:translateX(-50%) translateY(0)}
.count{font-size:12px;color:#64748b;margin-left:auto}
.more{text-align:center;padding:22px 0}
.more button{padding:11px 34px;border-radius:9px;border:1px solid #c7d2fe;
 background:#fff;color:#4338ca;font-size:13.5px;font-weight:600;cursor:pointer;font-family:inherit}
.more button:hover{background:#eef2ff}
.more span{display:block;font-size:12px;color:#94a3b8;margin-top:8px}
.flow2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
@media(max-width:700px){.flow2{grid-template-columns:1fr}}
.newscard{background:#fff;border-radius:10px;overflow:hidden;position:relative;
 box-shadow:0 1px 2px rgba(16,24,40,.07);display:flex;flex-direction:column}
.newscard.off{opacity:.42}
.chk{position:absolute;top:7px;left:7px;z-index:3;background:rgba(255,255,255,.92);
 border-radius:5px;padding:2px 5px;cursor:pointer}
.chk input{width:15px;height:15px;cursor:pointer;accent-color:#4f46e5}
.cat3{display:flex;flex-direction:column;gap:3px;margin-top:3px}
.ci{padding:4px 7px;border:1px solid #e2e8f0;border-radius:5px;font-size:11px;
 outline:none;font-family:inherit;width:100%;background:#fff;cursor:pointer}
select.ci{appearance:auto}
.ci:focus{border-color:#4f46e5;background:#fafaff}
.ci.ok{border-color:#10b981;background:#f0fdf4}
.simbtn{padding:4px 10px;border-radius:6px;border:1px solid #c7d2fe;background:#eef2ff;
 color:#4338ca;font-size:10.5px;font-weight:600;cursor:pointer;font-family:inherit}
.simbtn:hover{background:#e0e7ff}
.savebtn{border-color:#fdba74;background:#fff7ed;color:#c2410c}
.spuq{padding:4px 7px;border-radius:6px;border:1px solid #a5b4fc;background:#eef2ff;
 color:#4338ca;font-size:10.5px;cursor:pointer;white-space:nowrap}
.spuq:hover{background:#e0e7ff}
.spuedit{display:flex;gap:4px;margin-top:5px;align-items:center}
.libspu{flex:1;min-width:0;font-size:10.5px;padding:3px 6px;border:1px solid #e2e8f0;
 border-radius:5px;font-family:ui-monospace,Consolas,monospace;outline:none;background:#fff}
.libspu:focus{border-color:#4f46e5;background:#f8faff}
.libsave{padding:3px 7px;border-radius:5px;border:1px solid #86efac;background:#f0fdf4;
 color:#15803d;font-size:10.5px;cursor:pointer}
.libsave:hover{background:#dcfce7}
.libadd{padding:3px 7px;border-radius:5px;border:1px solid #fcd34d;background:#fffbeb;
 color:#b45309;font-size:11px;font-weight:700;cursor:pointer;line-height:1}
.libadd:hover{background:#fef3c7}
.pn.names{font-size:10.5px;line-height:1.45;max-height:88px;overflow:auto}
.nmrow{padding:1px 0;border-bottom:1px dashed #f1f5f9}
.nmrow:last-child{border-bottom:none}
.sitelist{margin-top:3px}
.sitelist .l1688{margin-top:2px}
.savebtn:hover{background:#ffedd5}
.links{margin:4px 0 6px}
.l1688{display:flex;gap:4px;margin-top:3px;align-items:center}
.i1688{flex:1;min-width:0;font-size:10.5px;padding:3px 5px;border:1px solid #e2e8f0;
  border-radius:5px;background:#f8fafc;color:#475569;font-family:ui-monospace,monospace;
  cursor:pointer;text-overflow:ellipsis}
.i1688:focus{background:#eff6ff;border-color:#93c5fd;outline:none}
.l1688 a{font-size:10.5px;color:#2563eb;text-decoration:none;white-space:nowrap}
.l1688 a:hover{text-decoration:underline}
.nolink{font-size:10.5px;color:#cbd5e1}
.siteurl{font-size:10.5px;color:#0f766e;text-decoration:none;background:#f0fdfa;
  border:1px solid #99f6e4;border-radius:5px;padding:3px 8px;max-width:100%;
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:inline-block}
.siteurl:hover{background:#ccfbf1;text-decoration:none}
.cnbadge{display:inline-block;font-size:10.5px;font-weight:600;color:#1d4ed8;
  background:#dbeafe;border:1px solid #bfdbfe;border-radius:4px;padding:1px 6px;
  vertical-align:1px;margin-left:2px;cursor:pointer}
.cnbadge:hover{background:#bfdbfe;border-color:#93c5fd}
.simbtn.done{background:#d1fae5;border-color:#6ee7b7;color:#047857}
.modal{position:fixed;inset:0;background:rgba(15,23,42,.62);z-index:200;
 display:none;align-items:center;justify-content:center;padding:20px}
.modal.on{display:flex}
.mbox{background:#fff;border-radius:14px;max-width:1080px;width:100%;max-height:92vh;
 overflow:auto;box-shadow:0 20px 60px rgba(0,0,0,.3)}
.mhead{padding:18px 24px;border-bottom:1px solid #e5e8ee;display:flex;
 align-items:center;gap:12px;position:sticky;top:0;background:#fff;z-index:2}
.mhead h3{font-size:16px;font-weight:700}
.mhead .close{margin-left:auto;border:none;background:#f1f5f9;width:30px;height:30px;
 border-radius:8px;cursor:pointer;font-size:16px;color:#64748b}
.mhead .close:hover{background:#e2e8f0}
.mbody{padding:20px 24px 26px}
.cmp{display:grid;grid-template-columns:200px 1fr 1fr;gap:14px;align-items:start;
 padding:14px 0;border-bottom:1px solid #f0f2f5}
.cmp:last-child{border-bottom:none}
.cmpimg{height:130px;background:#f8fafc;border-radius:8px;display:flex;
 align-items:center;justify-content:center;overflow:hidden}
.cmpimg img{width:100%;height:100%;object-fit:contain;background:#fff}
.cmpinfo{font-size:12.3px;color:#475569;display:flex;flex-direction:column;gap:5px}
.cmpinfo .cn{font-weight:600;color:#1e293b;line-height:1.4;
 display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
.cmpinfo .cs{display:flex;flex-wrap:wrap;gap:4px}
.level{display:inline-block;font-size:10.5px;padding:2px 8px;border-radius:5px;font-weight:700}
.level.same{background:#fee2e2;color:#b91c1c}
.level.likely{background:#fef3c7;color:#b45309}
.level.possible{background:#e0e7ff;color:#4338ca}
.cmpact{display:flex;flex-direction:column;gap:7px;justify-content:center}
.cmpact button{padding:7px 14px;border-radius:7px;border:none;font-size:12px;
 font-weight:600;cursor:pointer;font-family:inherit}
.byes{background:#10b981;color:#fff}
.byes:hover{background:#059669}
.bno{background:#fff;border:1px solid #e2e8f0;color:#64748b}
.bno:hover{background:#f8fafc}
.tgt{display:flex;gap:14px;padding:14px 16px;background:#f8fafc;border-radius:10px;
 margin-bottom:16px;align-items:center}
.tgt .cmpimg{height:110px;width:150px;flex:0 0 150px}
.simempty{text-align:center;padding:40px;color:#94a3b8;font-size:13.5px}
.doc{background:#fff;border-radius:12px;padding:28px 32px;max-width:900px;
 box-shadow:0 1px 3px rgba(16,24,40,.08)}
.doc h2{font-size:18px;margin:0 0 12px}
.doc ol,.doc ul{padding-left:22px;color:#475569;font-size:13.4px;line-height:2}
.doc p{color:#475569;font-size:13.4px;line-height:2;margin-top:8px}
.doc code{background:#f1f5f9;padding:1px 6px;border-radius:4px;
 font-family:ui-monospace,Consolas,monospace;font-size:12px}
</style></head>
<body>
<header><div class="hrow">
  <h1>Mico 选品工作台</h1>
  <span class="user">{{USER}} <a href="/logout" class="logout">退出</a></span>
  <div class="stats">
    <div class="st"><b>{{TOTAL}}</b><span>已询盘</span></div>
    <div class="st"><b>{{NSPU}}</b><span>有 SPU</span></div>
    <div class="st"><b>{{NMULTI}}</b><span>多 SPU</span></div>
    <div class="st"><b>{{NNOIQ}}</b><span>无询盘号</span></div>
    <div class="st"><b>{{NPAIRS}}</b><span>待配对</span></div>
  </div>
</div></header>
<main>
<div class="tabs">
  <div class="tab on" data-p="0" onclick="sw(0)">产品库 ({{TOTAL}})</div>
  <div class="tab" data-p="1" onclick="sw(1)">待采集配对 ({{NPAIRS}})</div>
  <div class="tab" data-p="3" onclick="sw(3)">未询盘新品 ({{NNEWS}})</div>
  <div class="tab" data-p="2" onclick="sw(2)">工作流说明</div>
</div>

<div class="panel on" id="p0">
  <div class="bar">
    <input class="search" id="q0" placeholder="搜索产品名称 / SPU / 询盘号…" oninput="f0(1)">
    <div class="chip on" data-f="all" onclick="cf(this)">全部</div>
    <div class="chip" data-f="multi" onclick="cf(this)">多 SPU ({{NMULTI}})</div>
    <div class="chip" data-f="nospu" onclick="cf(this)">无 SPU ({{NNOIQ}})</div>
    <div class="chip" data-f="noiq" onclick="cf(this)">无询盘号 ({{NNOIQ}})</div>
    <span class="count" id="c0"></span>
  </div>
  <div class="grid" id="g0">{{CARDS}}</div>
  <div class="more" id="more0"><button onclick="loadMore()">加载更多</button><button onclick="showAll0()">全部显示</button><span id="tip0"></span></div>
</div>

<div class="panel" id="p1">
  <div class="flow2">
    <div class="drop" id="dropq" onclick="document.getElementById('fq').click()">
      <div class="ico">📥</div>
      <h3>① 上传昨日询盘导出</h3>
      <p>template_center_inquiry_quote-export_*.xlsx</p>
      <input type="file" id="fq" accept=".xlsx" style="display:none" onchange="upq(this.files[0])">
    </div>
    <div class="drop" id="drop" onclick="document.getElementById('fi').click()">
      <div class="ico">📁</div>
      <h3>① 上传今日订单</h3>
      <p>order_task_result_*.xlsx</p>
      <input type="file" id="fi" accept=".xlsx" style="display:none" onchange="up(this.files[0])">
    </div>
  </div>
  <div class="summary" id="sum">
    <div class="srow" id="sumrow"></div>
    <div style="display:flex;gap:10px;flex-wrap:wrap">
      <button class="btn" onclick="saveSpuAll()">保存全部 SPU（云端）</button>
      <button class="btn ghost" onclick="exportXlsx()">导出配对表 xlsx</button>
    </div>
  </div>
  <div class="bar" style="margin-top:14px">
    <input class="search" id="q1" placeholder="搜索名称 / 询盘号…" oninput="f1()">
    <span class="count" id="c1"></span>
  </div>
  <div class="grid" id="g1">{{PCARDS}}</div>
  {{EMPTY}}
</div>

<div class="panel" id="p3">
  <div class="bar">
    <input class="search" id="q3" placeholder="搜索产品名称…" oninput="f3()">
    <button class="btn ghost" onclick="checkAll(true)">全选</button>
    <button class="btn ghost" onclick="checkAll(false)">全不选</button>
    <button class="btn" onclick="exportDemand()">导出询盘需求表</button>
    <span class="count" id="c3"></span>
  </div>
  <div class="grid" id="g3">{{NCARDS}}</div>
  {{NONEWS}}
</div>

<div class="panel" id="p2"><div class="doc">
<h2>每日工作流程</h2>
<ol>
<li>上传今日订单 xlsx → 自动识别「已询盘但无 SKU」的商品</li>
<li>在配对表里逐个填写 SPU</li>
<li>点「保存 SPU 并写回四列表」→ 自动备份原文件</li>
<li>点「导出配对表 xlsx」→ 得到当天配对表（含图片）</li>
</ol>
<h2 style="margin-top:26px">核心规则</h2>
<p><b>SPU 只认 SKU 列</b>：取 SKU 前缀，如 <code>ABC-blue-L</code> → <code>ABC</code>。
多个 SPU 会并列显示并标注「多」。</p>
<p><b>产品名称与客户标题一字不差</b>：不翻译、不改写、不加规格。</p>
<p><b>同款不同标题全部保留</b>：多站点多语言标题换行合并。</p>
<p><b>已询盘的不进需求表</b>；<b>订单无 SKU 的都是待配对</b>。</p>
<p><b>harpe.pl 店铺剔除</b>。</p>
</div></div>
</main>

<div class="modal" id="modal">
  <div class="mbox">
    <div class="mhead">
      <h3>相似款比对</h3>
      <button class="close" onclick="closeModal()">×</button>
    </div>
    <div class="mbody" id="mbody"></div>
  </div>
</div>

<div class="toast" id="toast"></div>
<script>
var FILTER='all', PAGE=1, PER={{PER}}, ALL=null, Q='';
function sw(i){
  document.querySelectorAll('.tab').forEach(function(t){
    t.classList.toggle('on', t.getAttribute('data-p')===String(i));
  });
  document.querySelectorAll('.panel').forEach(function(p){
    p.classList.toggle('on', p.id==='p'+i);
  });
  if(i===0)f0(0); if(i===1)f1(); if(i===3)f3();
}
function cf(el){
  document.querySelectorAll('.chip').forEach(function(c){c.classList.remove('on')});
  el.classList.add('on'); FILTER=el.dataset.f; f0(1);
}
function match(c){
  if(Q && (c.dataset.n.indexOf(Q)<0 && c.dataset.s.indexOf(Q)<0 && c.dataset.i.indexOf(Q)<0))return false;
  if(FILTER==='multi')return c.dataset.m==='1';
  if(FILTER==='nospu')return c.dataset.p==='1';
  if(FILTER==='noiq')return c.dataset.q==='1';
  return true;
}
function render(){
  var g=document.getElementById('g0');
  if(!ALL)ALL=Array.prototype.slice.call(g.children);
  var q=document.getElementById('q0').value.toLowerCase().trim();
  Q=q;
  var hits=ALL.filter(match);
  var show=hits.slice(0,PER);
  if(show.length<ALL.length || hits.length<ALL.length){
    g.innerHTML='';
    var frag=document.createDocumentFragment();
    show.forEach(function(c){frag.appendChild(c)});
    g.appendChild(frag);
  }
  document.getElementById('c0').textContent='显示 '+show.length+' / '+hits.length+' 款';
  var m=document.getElementById('more0');
  if(show.length<hits.length){
    m.style.display='';
    document.getElementById('tip0').textContent='还有 '+(hits.length-show.length)+' 款';
  }else{
    m.style.display='none';
  }
}
function f0(reset){ if(reset)PAGE=1; render(); }
function loadMore(){
  PAGE++;
  var g=document.getElementById('g0');
  var hits=ALL.filter(match);
  var more=hits.slice(g.children.length, g.children.length+PER*PAGE);
  var frag=document.createDocumentFragment();
  more.forEach(function(c){frag.appendChild(c)});
  g.appendChild(frag);
  document.getElementById('c0').textContent='显示 '+g.children.length+' / '+hits.length+' 款';
  var m=document.getElementById('more0');
  if(g.children.length<hits.length){
    document.getElementById('tip0').textContent='还有 '+(hits.length-g.children.length)+' 款';
  }else{
    m.style.display='none';
  }
}
function showAll0(){
  var g=document.getElementById('g0');
  var hits=ALL.filter(match);
  var frag=document.createDocumentFragment();
  hits.slice(g.children.length).forEach(function(c){frag.appendChild(c)});
  g.appendChild(frag);
  document.getElementById('more0').style.display='none';
  document.getElementById('c0').textContent='已显示全部 '+hits.length+' 款';
}
function f1(){
  var q=document.getElementById('q1').value.toLowerCase();
  var n=0;
  document.querySelectorAll('#g1 .paircard').forEach(function(c){
    var ok=!q || c.textContent.toLowerCase().indexOf(q)>-1;
    c.style.display=ok?'':'none'; if(ok)n++;
  });
  document.getElementById('c1').textContent='显示 '+n+' 款';
}
function toast(m){
  var t=document.getElementById('toast'); t.textContent=m;
  t.classList.add('show'); setTimeout(function(){t.classList.remove('show')},2600);
}
function byCn(cn){
  var q=document.getElementById('q0');
  if(q)q.value=cn;
  sw(0);
  window.scrollTo(0,0);
}
var drop=document.getElementById('drop');
if(drop){
  ['dragover','dragenter'].forEach(function(e){
    drop.addEventListener(e,function(ev){ev.preventDefault();drop.classList.add('over')})});
  ['dragleave','drop'].forEach(function(e){
    drop.addEventListener(e,function(ev){ev.preventDefault();drop.classList.remove('over')})});
  drop.addEventListener('drop',function(ev){
    if(ev.dataTransfer.files.length)up(ev.dataTransfer.files[0]);
  });
}
function up(file){
  if(!file)return;
  if(!/\.xlsx$/i.test(file.name)){toast('请上传 .xlsx 文件');return}
  var fd=new FormData(); fd.append('file',file);
  toast('正在处理 '+file.name+' …');
  fetch('/api/upload',{method:'POST',body:fd}).then(function(r){return r.json()})
  .then(function(d){
    if(!d.ok){toast('失败：'+(d.msg||'未知错误'));return}
    document.getElementById('sumrow').innerHTML=
      '<div class="si"><b>'+d.total+'</b>订单商品</div>'+
      '<div class="si"><b>'+d.inq+'</b>已询盘</div>'+
      '<div class="si"><b>'+d.with_sku+'</b>有 SKU</div>'+
      '<div class="si"><b>'+d.pairs+'</b>待配对</div>'+
      '<div class="si"><b>'+d.not_inq+'</b>未询盘</div>';
    document.getElementById('sum').classList.add('show');
    toast('已生成 '+d.pairs+' 款待配对');
    setTimeout(function(){location.reload()},1100);
  }).catch(function(e){toast('出错：'+e)});
}
document.addEventListener('input',function(e){
  if(e.target.classList&&e.target.classList.contains('spin')){
    e.target.classList.toggle('filled',e.target.value.trim().length>0);
  }
});
function collect(){
  var arr=[];
  document.querySelectorAll('.spin').forEach(function(i){
    arr.push({idx:parseInt(i.dataset.idx),spu:i.value.trim()});
  });
  return arr.filter(function(x){return x.spu});
}
var SAVING=false;
function doSave(arr,done,retry){
  if(!arr.length){toast('还没有填写任何 SPU');return}
  if(SAVING){toast('正在保存中…');return}
  SAVING=true;
  retry=retry||0;
  fetch('/api/save-spu',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({spus:arr})}).then(function(r){
      if(!r.ok)throw Error('HTTP '+r.status);
      return r.json()})
  .then(function(d){
    SAVING=false;
    if(!d.ok){
      toast('云端未写入：'+(d.msg||(d.errors&&d.errors[0])||'未知错误'));
      console.log('[save-spu]',d); return;
    }
    var m='已同步到云端 '+d.cloud+' 条';
    if(d.failed)m+='，失败 '+d.failed+' 条';
    if(d.xlsx_msg)m+='（'+d.xlsx_msg+'）';
    toast(m);
    arr.forEach(function(x){ if(x.spu) reflectLib(pairInq(x.idx),x.spu); });
    if(done)setTimeout(function(){location.reload()},1200);
  }).catch(function(e){
    SAVING=false;
    if(retry<2){
      toast('服务唤醒中，重试 '+(retry+1)+'/2 …');
      setTimeout(function(){doSave(arr,done,retry+1)},4000);
    }else{
      toast('保存失败：服务可能在冷启动，1 分钟后再试（'+e+'）');
    }
  });
}
function addPairSpu(idx){
  var el=document.querySelector('.spin[data-idx="'+idx+'"]');
  if(!el)return;
  var cur=el.value.trim();
  var add=prompt('追加 SPU（已有：'+(cur||'无')+'）');
  if(!add||!add.trim())return;
  var list=cur?cur.split(/[,\s/;]+/).filter(Boolean):[];
  add.split(/[,\s/;]+/).filter(Boolean).forEach(function(s){
    if(list.indexOf(s)<0)list.push(s);
  });
  el.value=list.join(',');
  el.classList.add('filled');
  doSave([{idx:idx,spu:el.value}],false);
}
function addNewsSpu(idx){
  var el=document.querySelector('.spn[data-idx="'+idx+'"]');
  if(!el)return;
  var cur=el.value.trim();
  var add=prompt('追加 SPU（已有：'+(cur||'无')+'）');
  if(!add||!add.trim())return;
  var list=cur?cur.split(/[,\s/;]+/).filter(Boolean):[];
  add.split(/[,\s/;]+/).filter(Boolean).forEach(function(s){
    if(list.indexOf(s)<0)list.push(s);
  });
  el.value=list.join(',');
  el.classList.add('filled');
  saveNewsSpu(idx);
}
function saveLibSpu(btn){
  var box=btn.parentNode;
  var inp=box.querySelector('.libspu');
  if(!inp)return;
  var inq=inp.getAttribute('data-inq')||'';
  var row=inp.getAttribute('data-row')||'';
  var v=inp.value.trim();
  fetch('/api/save-lib-spu',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({inq:inq,row:row,spu:v})}).then(function(r){
      if(!r.ok)throw Error('HTTP '+r.status);
      return r.json()}).then(function(d){
    if(!d.ok){toast('保存失败：'+(d.msg||''));return}
    var card=btn.closest('.pc');
    if(card){
      var ps=card.querySelector('.ps');
      if(ps)ps.innerHTML=(d.count>1
        ? d.spu.split(',').map(function(s){return '<b class="spu multi">'+esc2(s)+'</b>'}).join('')+'<b class="tag">多</b>'
        : (d.spu?'<b class="spu">'+esc2(d.spu)+'</b>':'<b class="spu none">无 SPU</b>'));
      card.dataset.s=(d.spu||'').toLowerCase();
      card.dataset.m=d.count>1?'1':'0';
      card.dataset.p=d.spu?'0':'1';
    }
    toast('产品库 SPU 已更新（'+(d.count||0)+' 个'+(d.merged>1?'，合并 '+(d.merged-1)+' 条同名产品':'')+'）');
  }).catch(function(e){toast('出错：'+e)});
}
function addLibSpu(btn){
  var box=btn.parentNode;
  var inp=box.querySelector('.libspu');
  if(!inp)return;
  var add=prompt('追加 SPU（已有：'+(inp.value.trim()||'无')+'）');
  if(!add||!add.trim())return;
  var cur=inp.value.trim();
  var list=cur?cur.split(/[,\s/;]+/).filter(Boolean):[];
  add.split(/[,\s/;]+/).filter(Boolean).forEach(function(s){
    if(list.indexOf(s)<0)list.push(s);
  });
  inp.value=list.join(',');
  saveLibSpu(btn);
}
function saveLibSpuLegacy(btn){}
function spuNames(idx){
  var el=document.querySelector('.spin[data-idx="'+idx+'"]');
  if(!el){toast('找不到输入框');return}
  var spu=el.value.trim();
  if(!spu){toast('先填 SPU 再查');return}
  var m=document.getElementById('modal'), b=document.getElementById('mbody');
  b.innerHTML='<div class="simempty">正在查询 '+esc2(spu)+' 关联的产品名…</div>';
  m.classList.add('on');
  fetch('/api/spu-names?spu='+encodeURIComponent(spu)).then(function(r){return r.json()})
  .then(function(d){
    var its=(d.items||[]);
    var h='<div style="font-weight:700;margin-bottom:8px">SPU <span class="iq">'+esc2(spu)+'</span> 关联的产品名（'+its.length+'）</div>';
    if(!its.length){
      h+='<div class="simempty">还没有记录。保存一次 SPU 后就会记录当前产品的所有客户产品名。</div>';
    }else{
      its.forEach(function(it){
        h+='<div class="cmp" style="align-items:flex-start">'
          +'<div class="cmpinfo" style="flex:1">'
          +'<div class="cn">'+esc2(it.name||'')+'</div>'
          +'<div class="pl">'+(it.url?'<a href="'+esc2(it.url)+'" target="_blank">站点</a>':'')
          +'<span class="iq">'+esc2((it.inq||'').slice(-14))+'</span>'
          +'<button class="simbtn" onclick="copyTxt(this)" data-t="'+esc2(it.name||'')+'">复制名</button></div>'
          +'</div></div>';
      });
    }
    b.innerHTML=h;
  }).catch(function(e){b.innerHTML='<div class="simempty">出错：'+esc2(e)+'</div>'});
}
function copyTxt(el){
  var t=el.getAttribute('data-t')||'';
  try{navigator.clipboard.writeText(t);toast('已复制产品名');return}catch(e){}
  var ta=document.createElement('textarea'); ta.value=t; document.body.appendChild(ta);
  ta.select(); document.execCommand('copy'); document.body.removeChild(ta); toast('已复制产品名');
}
function groupIdxs(el,idx){
  var s=el&&el.getAttribute('data-idxs');
  if(!s)return [idx];
  return s.split(',').map(function(x){return parseInt(x,10)}).filter(function(x){return !isNaN(x)});
}
function saveOne(idx){
  var el=document.querySelector('.spin[data-idx="'+idx+'"]');
  if(!el){toast('找不到输入框');return}
  var v=el.value.trim();
  if(!v){toast('这条还没填 SPU');return}
  el.classList.add('filled');
  var arr=groupIdxs(el,idx).map(function(k){return {idx:k,spu:v}});
  doSave(arr,false);
}
function saveSpuAll(){
  saveDemand();
  var arr=collect();
  if(!arr.length){toast('还没有填写任何 SPU');return}
  if(!confirm('把 '+arr.length+' 条 SPU 保存到云端？（同事立即可见）'))return;
  doSave(arr,true);
}
function saveSpu(){
  saveDemand();
  var arr=collect();
  if(!arr.length){toast('还没有填写任何 SPU');return}
  if(!confirm('确定把 '+arr.length+' 条 SPU 写回四列表？\n（会自动备份原文件）'))return;
  fetch('/api/save-spu',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({spus:arr})}).then(function(r){return r.json()})
  .then(function(d){
    if(!d.ok){toast('失败：'+(d.msg||''));return}
    toast('已更新 '+d.updated+' 条，备份：'+(d.backup||'无'));
  }).catch(function(e){toast('出错：'+e)});
}
function saveNewsSpu(idx){
  var el=document.querySelector('.spn[data-idx="'+idx+'"]');
  if(!el){toast('未找到输入框');return}
  var v=el.value.trim();
  if(!v){toast('这条还没填 SPU');return}
  var items=groupIdxs(el,idx).map(function(k){return {idx:k,spu:v}});
  fetch('/api/save-news-spu',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({items:items})}).then(function(r){
      if(!r.ok)throw Error('HTTP '+r.status);
      return r.json()})
  .then(function(d){
    if(!d.ok){toast('写入失败：'+(d.errors&&d.errors[0]||'未知错误'));return}
    el.classList.add('filled');
    if(d.moved){
      toast('SPU 已保存，'+d.moved+' 款已移入待采集配对');
      setTimeout(function(){location.reload()},1000);
    }else{
      toast('SPU 已保存到云端');
    }
  }).catch(function(e){toast('出错：'+e)});
}
function exportXlsx(){
  toast('正在导出，请稍候…');
  fetch('/api/export',{method:'POST'}).then(function(r){return r.json()})
  .then(function(d){
    if(!d.ok){toast('失败：'+(d.msg||''));return}
    toast('已导出 '+d.file+'（'+d.imgs+' 张图）');
  }).catch(function(e){toast('出错：'+e)});
}
function f3(){
  var q=document.getElementById('q3').value.toLowerCase();
  var n=0, tot=0;
  document.querySelectorAll('#g3 .newscard').forEach(function(c){
    tot++;
    var ok=!q || c.textContent.toLowerCase().indexOf(q)>-1;
    c.style.display=ok?'':'none'; if(ok)n++;
  });
  document.getElementById('c3').textContent='显示 '+n+' / '+tot+' 款';
}
function checkAll(v){
  document.querySelectorAll('#g3 .newscard').forEach(function(c,i){
    var cb=c.querySelector('input[type=checkbox]');
    if(cb){cb.checked=v; c.classList.toggle('off',!v);}
  });
}
function upq(file){
  if(!file)return;
  if(!/\.xlsx$/i.test(file.name)){toast('请上传 .xlsx 文件');return}
  var fd=new FormData(); fd.append('file',file);
  toast('正在导入询盘文件…');
  fetch('/api/upload-quote',{method:'POST',body:fd}).then(function(r){return r.json()})
  .then(function(d){
    if(!d.ok){toast('失败：'+(d.msg||''));return}
    var m='';
    if(d.pairs_added)m='已加入待采集配对 '+d.pairs_added+' 款';
    if(d.news_removed)m+=(m?'，':'')+'从新品列表剔除 '+d.news_removed+' 款';
    if(d.added>0)m='导入 '+d.added+' 款新品'+(m?'，'+m:'');
    if(!m)m='询盘文件 '+d.total+' 款均已处理过';
    toast(m);
    setTimeout(function(){location.reload()},1400);
  }).catch(function(e){toast('出错：'+e)});
}
document.addEventListener('input',function(e){
  if(e.target.classList&&e.target.classList.contains('ci')){
    e.target.classList.toggle('ok',e.target.value.trim().length>0);
  }
});
document.addEventListener('change',function(e){
  if(e.target.type==='checkbox' && e.target.closest('#g3')){
    var card=e.target.closest('.newscard');
    if(card)card.classList.toggle('off',!e.target.checked);
  }
});
function collectDemand(){
  var arr=[];
  document.querySelectorAll('#g3 .newscard').forEach(function(c){
    var idx=parseInt(c.dataset.idx);
    var cb=c.querySelector('input[type=checkbox]');
    var l1=c.querySelector('.ci[data-k=l1]');
    var l2=c.querySelector('.ci[data-k=l2]');
    var l3=c.querySelector('.ci[data-k=l3]');
    arr.push({idx:idx, checked:cb?cb.checked:true,
      l1:l1?l1.value.trim():'', l2:l2?l2.value.trim():'', l3:l3?l3.value.trim():''});
  });
  return arr;
}
function saveDemand(){
  return fetch('/api/save-demand',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({items:collectDemand()})}).then(function(r){return r.json()});
}
function exportDemand(){
  var items=collectDemand();
  var n=items.filter(function(x){return x.checked}).length;
  if(!n){toast('请至少勾选 1 款');return}
  saveDemand().then(function(){
    return fetch('/api/export-demand',{method:'POST'});
  }).then(function(r){
    var ct=r.headers.get('Content-Type')||'';
    if(ct.indexOf('spreadsheet')>=0){
      return r.blob().then(function(bl){
        var cd=r.headers.get('Content-Disposition')||'';
        var m=/filename\*=UTF-8''([^;]+)/.exec(cd);
        var name=m?decodeURIComponent(m[1]):'询盘需求表.xlsx';
        var a=document.createElement('a');
        a.href=URL.createObjectURL(bl);
        a.download=name;
        document.body.appendChild(a); a.click();
        setTimeout(function(){URL.revokeObjectURL(a.href);a.remove()},1500);
        toast('已下载 '+name);
      });
    }
    return r.json().then(function(d){toast('失败：'+(d.msg||'未知'))});
  }).catch(function(e){toast('出错：'+e)});
}
var SIMKIND='pair', SIMIDX=-1;
function catL1(sel){
  var card=sel.closest('.cat3');
  var l2=card.querySelector('[data-k=l2]');
  var l3=card.querySelector('[data-k=l3]');
  l3.innerHTML='<option value="">三级分类</option>';
  if(!sel.value){l2.innerHTML='<option value="">二级分类</option>';return}
  fetch('/api/cat-l2?l1='+encodeURIComponent(sel.value))
    .then(function(r){return r.json()}).then(function(d){
      var h='<option value="">二级分类</option>';
      (d.items||[]).forEach(function(v){
        h+='<option value="'+esc2(v)+'">'+esc2(v)+'</option>';
      });
      l2.innerHTML=h;
    });
}
function catL2(sel){
  var card=sel.closest('.cat3');
  var l1=card.querySelector('[data-k=l1]');
  var l3=card.querySelector('[data-k=l3]');
  l3.innerHTML='<option value="">三级分类</option>';
  if(!sel.value||!l1.value)return;
  fetch('/api/cat-l3?l1='+encodeURIComponent(l1.value)+'&l2='+encodeURIComponent(sel.value))
    .then(function(r){return r.json()}).then(function(d){
      var h='<option value="">三级分类</option>';
      (d.items||[]).forEach(function(v){
        h+='<option value="'+esc2(v)+'">'+esc2(v)+'</option>';
      });
      l3.innerHTML=h;
    });
}
function closeModal(){document.getElementById('modal').classList.remove('on')}
function esc2(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}
function spuHtml(c){
  if(!c.spus || !c.spus.length) return '<span class="spu none">无 SPU</span>';
  var h='';
  c.spus.forEach(function(s){
    h+='<b class="spu'+(c.spus.length>1?' multi':'')+'">'+esc2(s)+'</b>';
  });
  if(c.spus.length>1)h+='<b class="tag">多</b>';
  return h;
}
function sim(idx,kind){
  SIMKIND=kind; SIMIDX=idx;
  var m=document.getElementById('modal');
  var b=document.getElementById('mbody');
  b.innerHTML='<div class="simempty">正在比对，首次需下载图片，请稍候…</div>';
  m.classList.add('on');
  fetch('/api/similar',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({idx:idx,kind:kind})}).then(function(r){return r.json()})
  .then(function(d){
    if(!d.ok){b.innerHTML='<div class="simempty">'+esc2(d.msg||'比対失败')+'</div>';return}
    var t=d.target, cs=d.candidates;
    var h='<div class="tgt"><div class="cmpimg">';
    h+= t.img? '<img src="'+esc2(t.img)+'">' : '<span class="nopic">无图</span>';
    h+='</div><div class="cmpinfo"><div style="font-size:11px;color:#94a3b8">比对目标</div>';
    h+='<div class="cn" style="font-weight:700">'+esc2(t.names[0]||'')+(t.cn?' <b class="cnbadge" onclick="byCn(\''+esc2(t.cn)+'\')">'+esc2(t.cn)+'</b>':'')+'</div></div></div>';
    if(!cs.length){
      h+='<div class="simempty">已询盘库中<b>没有找到相似产品</b><br>可以放心当作新品处理</div>';
    }else{
      h+='<div style="font-size:12.5px;color:#64748b;margin-bottom:6px">找到 '
         +cs.length+' 个相似款，请确认是否为同款：</div>';
      cs.forEach(function(c,k){
        h+='<div class="cmp">';
        h+='<div class="cmpimg">'+(c.img?'<img src="'+esc2(c.img)+'">':'<span class="nopic">无图</span>')+'</div>';
        h+='<div class="cmpinfo">';
        h+='<div class="cn">'+esc2(c.name)+(c.cn?' <b class="cnbadge" onclick="byCn(\''+esc2(c.cn)+'\')">'+esc2(c.cn)+'</b>':'')+'</div>';
        h+='<div class="cs">'+spuHtml(c)+'</div>';
        h+='<div style="font-size:11px;color:#94a3b8">'
           +(c.dist!==null?('图距 '+c.dist+' · '):'')
           +'标题相似 '+c.nsim+'</div>';
        h+='<div class="pl"><a href="'+esc2(c.url)+'" target="_blank">站点</a>'
           +'<span class="iq">'+esc2((c.inq||'').slice(-14))+'</span></div>';
        h+='</div>';
        h+='<div class="cmpact">';
        h+='<span class="level '+c.level+'">'+esc2(c.level_txt)+'</span>';
        h+='<button class="byes" data-spu="'+esc2(c.spu||(c.spus||[]).join(','))+'" '
           +'onclick="markSame(this,'+idx+',\''+kind+'\','+c.row+')">是同一款</button>';
        h+='<button class="bno" onclick="markDiff('+idx+',\''+kind+'\')">不是同款</button>';
        h+='</div></div>';
      });
    }
    b.innerHTML=h;
  }).catch(function(e){b.innerHTML='<div class="simempty">出错：'+esc2(e)+'</div>'});
}
function reflectLib(inq,spu){
  var pc=document.querySelector('#g0 .pc[data-i="'+String(inq||'').toLowerCase()+'"]');
  if(!pc)return;
  var ps=pc.querySelector('.ps'); if(!ps)return;
  if(spu){
    ps.innerHTML='<b class="spu">'+esc2(spu)+'</b>';
    pc.dataset.s=String(spu).toLowerCase(); pc.dataset.p='0';
  }
}
function pairInq(idx){
  var el=document.querySelector('.spin[data-idx="'+idx+'"]');
  var card=el&&el.closest('.paircard');
  return card?card.getAttribute('data-inq'):'';
}
function markSame(btn,idx,kind,row){
  var spu=(btn.getAttribute('data-spu')||'').trim();
  fetch('/api/mark-same',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({idx:idx,kind:kind,same:true,cand_row:row})})
  .then(function(r){return r.json()}).then(function(d){
    closeModal();
    if(!d.ok){toast('标记保存失败：'+(d.msg||''));return}
    if(kind==='news'&&d.removed){
      toast('库里已有同款，已移出未询盘新品');
      setTimeout(function(){location.reload()},900);return;
    }
    if(kind==='pair'&&spu){
      var el=document.querySelector('.spin[data-idx="'+idx+'"]');
      if(el){el.value=spu;el.classList.add('filled');}
      doSave([{idx:idx,spu:spu}],false);
      toast('已标同款，SPU '+spu+' 已带出并保存');
    }else{
      toast('已标记为同款，可据此剔除');
    }
  }).catch(function(e){toast('出错：'+e)});
}
function markDiff(idx,kind){
  fetch('/api/mark-same',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({idx:idx,kind:kind,same:false})})
  .then(function(r){return r.json()}).then(function(d){
    toast(d.ok?'已标记为不同款':'保存失败');
    closeModal();
  }).catch(function(e){toast('出错：'+e)});
}
document.getElementById('modal').addEventListener('click',function(e){
  if(e.target===this)closeModal();
});
f0(0); f1(); f3();
</script></body></html>"""


# ================= 服务 =================
class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, *a):
        pass

    # ---------- 工具 ----------
    def _send(self, code, body, ctype="application/json; charset=utf-8", extra=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        if extra:
            for k, v in extra:
                self.send_header(k, v)
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _redirect(self, url, extra=None):
        self.send_response(302)
        self.send_header("Location", url)
        self.send_header("Content-Length", "0")
        self.send_header("Cache-Control", "no-store")
        if extra:
            for k, v in extra:
                self.send_header(k, v)
        self.end_headers()

    def _user(self):
        """返回当前登录用户名，未登录返回 None。"""
        ck = parse_cookie(self.headers.get("Cookie", ""))
        u = unquote(ck.get("user", ""))
        t = ck.get(COOKIE_NAME, "")
        if u and APP_PASSWORD and t == make_token(u):
            return u
        if not APP_PASSWORD:
            return u or "本机"
        return None

    def _need_auth(self):
        u = self._user()
        if u is None:
            if self.path.startswith("/api/"):
                self._send(401, json.dumps({"ok": False, "msg": "未登录"}, ensure_ascii=False))
            else:
                self._redirect("/login")
            return None
        CURRENT_USER["name"] = u
        return u

    # ---------- GET ----------
    def do_GET(self):
        try:
            if self.path == "/login":
                self._send(200, page_login(), "text/html; charset=utf-8")
                return
            if self.path == "/logout":
                self._redirect("/login", [("Set-Cookie",
                    "user=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax"),
                    ("Set-Cookie", "%s=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax" % COOKIE_NAME)])
                return

            if not self._need_auth():
                return

            if self.path in ("/", "/index"):
                self._send(200, page_index(), "text/html; charset=utf-8")
            elif self.path == "/api/refresh":
                self._send(200, json.dumps(sync_from_excel(), ensure_ascii=False))
            elif self.path.startswith("/api/spu-names"):
                from urllib.parse import unquote
                q = unquote(self.path.split("?", 1)[1]) if "?" in self.path else ""
                spu = ""
                for kv in q.split("&"):
                    if kv.startswith("spu="):
                        spu = kv[4:].strip()
                rows = sb.get_spu_names(spu) if sb.enabled() else []
                out = [{"spu": r.get("spu"), "name": r.get("name"),
                        "url": r.get("url"), "inq": r.get("inq")} for r in rows]
                self._send(200, json.dumps({"ok": True, "spu": spu, "items": out},
                                           ensure_ascii=False))
            elif self.path.startswith("/api/cat-l2"):
                q = parse_qs(urlparse(self.path).query)
                l1 = (q.get("l1") or [""])[0]
                items = sorted((cat_tree().get(l1) or {}).keys())
                self._send(200, json.dumps({"items": items}, ensure_ascii=False))
            elif self.path.startswith("/api/cat-l3"):
                q = parse_qs(urlparse(self.path).query)
                l1 = (q.get("l1") or [""])[0]
                l2 = (q.get("l2") or [""])[0]
                items = sorted(((cat_tree().get(l1) or {}).get(l2) or []))
                self._send(200, json.dumps({"items": items}, ensure_ascii=False))
            else:
                self._send(404, json.dumps({"ok": False}, ensure_ascii=False))
        except Exception as e:
            traceback.print_exc()
            self._send(500, json.dumps({"ok": False, "msg": str(e)}, ensure_ascii=False))

    # ---------- POST ----------
    def do_POST(self):
        try:
            n = int(self.headers.get("Content-Length", 0) or 0)
            raw = self.rfile.read(n) if n else b""

            # 登录
            if self.path == "/login":
                q = parse_qs(raw.decode("utf-8", "replace"))
                u = (q.get("user") or [""])[0]
                pw = (q.get("pwd") or [""])[0]
                if pw != APP_PASSWORD:
                    self._send(200, page_login("密码不正确"), "text/html; charset=utf-8")
                    return
                if u not in USERS:
                    self._send(200, page_login("请选择你的姓名"), "text/html; charset=utf-8")
                    return
                tok = make_token(u)
                self._redirect("/", [("Set-Cookie",
                    "user=%s; Path=/; Max-Age=2592000; HttpOnly; SameSite=Lax" % quote(u)),
                    ("Set-Cookie",
                     "%s=%s; Path=/; Max-Age=2592000; HttpOnly; SameSite=Lax" % (COOKIE_NAME, tok))])
                return

            user = self._need_auth()
            if not user:
                return
            CURRENT_USER["name"] = user

            # 文件上传
            if self.path in ("/api/upload", "/api/upload-quote"):
                ctype = self.headers.get("Content-Type", "")
                bnd = None
                for part in ctype.split(";"):
                    part = part.strip()
                    if part.startswith("boundary="):
                        bnd = part[9:].strip('"')
                if not bnd:
                    self._send(200, json.dumps({"ok": False, "msg": "上传格式错误"}, ensure_ascii=False))
                    return
                filedata = None
                for seg in raw.split(b"--" + bnd.encode()):
                    if b"filename=" in seg:
                        idx = seg.find(b"\r\n\r\n")
                        if idx > 0:
                            filedata = seg[idx + 4:].rsplit(b"\r\n", 1)[0]
                        break
                if not filedata:
                    self._send(200, json.dumps({"ok": False, "msg": "未收到文件"}, ensure_ascii=False))
                    return
                tag = "quote" if self.path.endswith("quote") else "order"
                tmp = os.path.join(DATA_DIR, "_upload_%s.xlsx" % tag)
                with open(tmp, "wb") as f:
                    f.write(filedata)
                try:
                    r = import_quotes(tmp) if tag == "quote" else process_order(tmp)
                except Exception as e:
                    traceback.print_exc()
                    r = {"ok": False, "msg": str(e)}
                try:
                    os.remove(tmp)
                except Exception:
                    pass
                self._send(200, json.dumps(r, ensure_ascii=False))
                return

            # 保存 SPU
            if self.path == "/api/save-spu":
                payload = json.loads(raw.decode("utf-8"))
                arr = payload.get("spus", [])
                d = load_data()
                pairs = d.get("pairs", [])
                for x in arr:
                    i2 = x.get("idx")
                    if isinstance(i2, int) and 0 <= i2 < len(pairs):
                        pairs[i2]["filled_spu"] = x.get("spu", "")
                        pairs[i2]["by"] = user
                d["pairs"] = pairs
                save_data(d)
                record_edit(pairs, user, arr)
                # 关键：先落 Supabase（多人共享），再尽力写本地 xlsx
                ok_n, fail_n, errs = save_spu_supabase(arr, user)
                res = {"ok": ok_n > 0 or not arr, "cloud": ok_n,
                       "failed": fail_n, "errors": errs[:5]}
                try:
                    xl = apply_spu_to_master(user, sync_pairs=False)
                    res["updated"] = xl.get("updated", 0)
                    res["backup"] = xl.get("backup")
                    if not xl.get("ok"):
                        res["xlsx_msg"] = xl.get("msg", "")
                except Exception as e:
                    res["xlsx_msg"] = "xlsx 写入失败（不影响云端）：" + str(e)[:80]
                if fail_n:
                    res["msg"] = "%d 条未写入云端" % fail_n
                self._send(200, json.dumps(res, ensure_ascii=False))
                return

            # 保存需求表
            if self.path == "/api/save-demand":
                payload = json.loads(raw.decode("utf-8"))
                d = load_data()
                news = d.get("news", [])
                for x in payload.get("items", []):
                    i2 = x.get("idx")
                    if isinstance(i2, int) and 0 <= i2 < len(news):
                        news[i2]["checked"] = bool(x.get("checked"))
                        news[i2]["l1"] = x.get("l1", "") or ""
                        news[i2]["l2"] = x.get("l2", "") or ""
                        news[i2]["l3"] = x.get("l3", "") or ""
                        news[i2]["by"] = user
                d["news"] = news
                save_data(d)
                if sb.enabled():
                    try:
                        for x in payload.get("items", []):
                            i2 = x.get("idx")
                            if isinstance(i2, int) and 0 <= i2 < len(news):
                                nid = news[i2].get("id")
                                if nid:
                                    sb.update_news(nid, {
                                        "checked": bool(x.get("checked")),
                                        "l1": x.get("l1", "") or None,
                                        "l2": x.get("l2", "") or None,
                                        "l3": x.get("l3", "") or None,
                                    }, user)
                    except Exception as e:
                        print("[Supabase 需求表同步失败] %s" % e)
                self._send(200, json.dumps({"ok": True}, ensure_ascii=False))
                return

            # 产品库：修改/补充 SPU（支持一个产品多个 SPU）
            if self.path == "/api/save-lib-spu":
                payload = json.loads(raw.decode("utf-8"))
                inq = (payload.get("inq") or "").strip()
                row = payload.get("row")
                raw_spu = (payload.get("spu") or "").strip()
                spus = sb.split_spu(raw_spu)
                spu_txt = ",".join(spus)
                d = load_data()
                target = None
                for p in d.get("products", []):
                    if inq and p.get("inq") == inq:
                        target = p
                        break
                    if row is not None and str(p.get("row")) == str(row):
                        target = p
                        break
                if target is None:
                    self._send(200, json.dumps(
                        {"ok": False, "msg": "产品库里没找到这条记录"}, ensure_ascii=False))
                    return
                # 合并组：同询盘号 / 同旧SPU 的所有产品一起更新（一个货号多站点）
                old_spus = [s for s in (target.get("spus") or []) if s]
                siblings = []
                for p in d.get("products", []):
                    if p is target:
                        siblings.append(p)
                        continue
                    same_inq = inq and p.get("inq") == inq
                    same_spu = old_spus and (set(p.get("spus") or []) & set(old_spus))
                    same_url = (p.get("url") or "") and (p.get("url") == target.get("url"))
                    if same_inq or same_spu or same_url:
                        siblings.append(p)
                for p in siblings:
                    p["spu"] = spu_txt
                    p["spus"] = spus
                    p["multi_spu"] = len(spus) > 1
                if sb.enabled():
                    try:
                        for p in siblings:
                            match = ({"inq": p["inq"]} if p.get("inq")
                                     else {"url_key": sb.norm_url(p.get("url") or "")})
                            sb.update("products", match, {"spu": spu_txt or None, "spus": spus})
                        sb.put_spu_names([{"spu": s, "name": n,
                                           "url": p.get("url") or "",
                                           "inq": p.get("inq") or ""}
                                          for p in siblings for s in spus
                                          for n in (p.get("names") or [])])
                        sb.log_edit(user, "products", inq or target.get("url"),
                                    "", spu_txt,
                                    "%s（合并 %d 条）" % (
                                        (target.get("names") or [""])[0][:60], len(siblings)))
                    except Exception as e:
                        self._send(200, json.dumps(
                            {"ok": False, "msg": "云端写入失败：%s" % str(e)[:120]},
                            ensure_ascii=False))
                        return
                save_data(d)
                invalidate_cache()
                self._send(200, json.dumps(
                    {"ok": True, "spu": spu_txt, "count": len(spus),
                     "merged": len(siblings)},
                    ensure_ascii=False))
                return

            if self.path == "/api/export":
                self._send(200, json.dumps(export_pairs(), ensure_ascii=False))
                return

            if self.path == "/api/export-demand":
                res = export_demand()
                if res.get("ok"):
                    fp = os.path.join(BASE, res["file"])
                    try:
                        with open(fp, "rb") as f:
                            data = f.read()
                        self._send(200, data,
                                   "application/vnd.openxmlformats-officedocument"
                                   ".spreadsheetml.sheet",
                                   [("Content-Disposition",
                                     "attachment; filename*=UTF-8''%s"
                                     % quote(res["file"]))])
                        return
                    except Exception as e:
                        res = {"ok": False, "msg": "读取文件失败：%s" % str(e)[:80]}
                self._send(200, json.dumps(res, ensure_ascii=False))
                return

            if self.path == "/api/similar":
                payload = json.loads(raw.decode("utf-8"))
                kind = payload.get("kind", "pair")
                i2 = payload.get("idx")
                d = load_data()
                if kind == "lib":
                    # 产品库内部互相比对：目标取自产品库，候选也是产品库
                    lst = d.get("products", [])
                else:
                    lst = d.get("pairs" if kind == "pair" else "news", [])
                if not (isinstance(i2, int) and 0 <= i2 < len(lst)):
                    self._send(200, json.dumps({"ok": False, "msg": "索引无效"}, ensure_ascii=False))
                    return
                t = lst[i2]
                names = t.get("names") or [t.get("name", "")]
                _sp = set()
                for z in (t.get("spus") or []):
                    if z:
                        _sp.add(z)
                if t.get("spu"):
                    _sp.add(t["spu"])
                cands = similar.find_similar(names, t.get("img", ""),
                                              d.get("products", []),
                                              exclude_inq=t.get("inq") or None,
                                              exclude_spu=_sp, t_url=t.get("url"))
                out = []
                for c in cands:
                    c["level_txt"] = similar.LEVEL_TXT.get(c["level"], "")
                    out.append(c)
                self._send(200, json.dumps(
                    {"ok": True, "target": {"names": names, "img": t.get("img", ""),
                                            "cn": similar.cn_for(t)},
                     "candidates": out}, ensure_ascii=False))
                return

            # 未询盘新品：手动保存 SPU → 有 SPU 后移入「待采集配对」
            if self.path == "/api/save-news-spu":
                payload = json.loads(raw.decode("utf-8"))
                arr = payload.get("items", [])
                d = load_data()
                news = d.get("news", [])
                pairs = d.get("pairs", [])
                ok_n = 0
                errs = []
                moved_ids = []
                for x in arr:
                    i2 = x.get("idx")
                    spu = (x.get("spu") or "").strip()
                    if not isinstance(i2, int) or not (0 <= i2 < len(news)):
                        continue
                    item = news[i2]
                    item["filled_spu"] = spu
                    item["by"] = user
                    ok_n += 1
                    if sb.enabled() and item.get("id"):
                        try:
                            sb.update("news", {"id": item["id"]},
                                      {"filled_spu": spu, "edited_by": user})
                            sb.log_edit(user, "news", item.get("id"), "", spu,
                                        (item.get("names") or [item.get("name", "")])[0][:80])
                        except Exception as e:
                            errs.append(str(e)[:120])
                            continue
                    # 有 SPU = 货源已确认 → 移入待采集配对 + 写进产品库
                    if spu:
                        if sb.enabled():
                            try:
                                sb.put_spu_names([{"spu": y.strip(), "name": n,
                                                   "url": item.get("url") or "",
                                                   "inq": item.get("inq") or ""}
                                                  for y in re.split(r"[,\s/;]+", spu) if y.strip()
                                                  for n in (item.get("names") or [item.get("name", "")])])
                            except Exception:
                                pass
                            # 写进产品库：以后导入订单按产品库匹配就能认出它已确定货源
                            try:
                                m0 = re.search(r"INQ-(\d+)", item.get("inq") or "")
                                sb.upsert_products([{
                                    "row": None,
                                    "names": item.get("names") or [item.get("name", "")],
                                    "url": item.get("url") or "",
                                    "inq": item.get("inq") or None,
                                    "batch": m0.group(1)[-4:] if m0 else None,
                                    "spu": spu, "spus": sb.split_spu(spu),
                                    "u1688": item.get("u1688") or [],
                                    "img": item.get("img") or None,
                                }])
                            except Exception as e:
                                errs.append("写产品库失败:%s" % str(e)[:80])
                        u = norm(item.get("url"))
                        dup = any(norm(p.get("url")) == u for p in pairs if p.get("url"))
                        if not dup and u:
                            pairs.append({
                                "url": item.get("url") or "",
                                "names": item.get("names") or [item.get("name", "")],
                                "inq": item.get("inq") or "",
                                "spu": spu, "spus": sb.split_spu(spu),
                                "u1688": item.get("u1688") or [],
                                "img": item.get("img") or "",
                                "filled_spu": spu, "mark": None, "note": "",
                                "l1": item.get("l1") or "", "l2": item.get("l2") or "",
                                "l3": item.get("l3") or "",
                            })
                        if sb.enabled() and item.get("id"):
                            try:
                                sb.delete("news", {"id": item["id"]})
                            except Exception as e:
                                errs.append("移出新品失败:%s" % str(e)[:90])
                                continue
                        if sb.enabled():
                            try:
                                sb.log_op(user, "news_spu_to_pairs",
                                          {"name": (item.get("names") or [item.get("name", "")])[0][:80],
                                           "spu": spu, "url": item.get("url") or ""})
                            except Exception:
                                pass
                        moved_ids.append(item.get("id"))
                if moved_ids:
                    news = [n for n in news if n.get("id") not in moved_ids]
                    d["news"] = news
                    d["pairs"] = pairs
                save_data(d)
                if sb.enabled() and moved_ids:
                    try:
                        sb.save_pairs(d.get("pairs_date") or datetime.date.today().isoformat(), pairs)
                    except Exception as e:
                        errs.append("配对写入失败:%s" % str(e)[:90])
                invalidate_cache()
                self._send(200, json.dumps(
                    {"ok": not errs, "saved": ok_n, "moved": len(moved_ids),
                     "errors": errs[:2]},
                    ensure_ascii=False))
                return

            if self.path == "/api/mark-same":
                payload = json.loads(raw.decode("utf-8"))
                kind = payload.get("kind", "pair")
                i2 = payload.get("idx")
                same = bool(payload.get("same"))
                cand_row = payload.get("cand_row")
                d = load_data()
                key = "pairs" if kind == "pair" else "news"
                lst = d.get(key, [])
                if isinstance(i2, int) and 0 <= i2 < len(lst):
                    # 新品确认同款 → 并入产品库（带库里同款的 SPU），并从新品移除
                    if kind == "news" and same:
                        item = lst[i2]
                        cand = None
                        rows = d.get("products", [])
                        for pr in rows:
                            if pr.get("row") == cand_row:
                                cand = pr
                                break
                        spu_val = (cand or {}).get("spu") or ""
                        # 并入产品库：作为库内同款的另一个客户名称
                        if cand:
                            try:
                                nm = (item.get("names") or [item.get("name", "")])
                                sb.upsert_products([{
                                    "row": None, "names": nm,
                                    "url": item.get("url") or "",
                                    "inq": cand.get("inq") or "",
                                    "batch": cand.get("batch") or "",
                                    "spu": spu_val,
                                    "spus": sb.split_spu(spu_val),
                                    "u1688": cand.get("u1688") or [],
                                    "img": item.get("img") or cand.get("img") or "",
                                }])
                                if spu_val:
                                    sb.put_spu_names([{"spu": y.strip(), "name": n,
                                                       "url": item.get("url") or "",
                                                       "inq": cand.get("inq") or ""}
                                                      for y in re.split(r"[,\s/;]+", spu_val) if y.strip()
                                                      for n in nm])
                                sb.log_op(user, "news_same_merged_to_library",
                                          {"name": (nm or [""])[0][:80], "spu": spu_val,
                                           "same_as_row": cand_row,
                                           "url": item.get("url") or ""})
                            except Exception as e:
                                print("[合并入库失败] %s" % e)
                        if sb.enabled():
                            if not item.get("id"):
                                self._send(200, json.dumps(
                                    {"ok": False, "msg": "该新品缺少云端ID，无法移除"},
                                    ensure_ascii=False))
                                return
                            try:
                                sb.delete("news", {"id": item["id"]})
                            except Exception as e:
                                self._send(200, json.dumps(
                                    {"ok": False, "msg": "云端移除失败：%s" % str(e)[:80]},
                                    ensure_ascii=False))
                                return
                        # 同款确认 = 自带 SPU → 同时写进待采集配对，保留这条的处理结果
                        pairs_now = d.get("pairs", [])
                        pu = norm(item.get("url"))
                        already = any(pu and norm(p.get("url")) == pu for p in pairs_now)
                        if not already:
                            pairs_now.append({
                                "url": item.get("url") or "",
                                "names": item.get("names") or [item.get("name", "")],
                                "inq": (cand or {}).get("inq") or "",
                                "spu": spu_val,
                                "u1688": (cand or {}).get("u1688") or item.get("u1688") or [],
                                "img": item.get("img") or (cand or {}).get("img") or "",
                                "filled_spu": spu_val, "mark": "same", "note": "",
                                "l1": item.get("l1") or "", "l2": item.get("l2") or "",
                                "l3": item.get("l3") or "",
                            })
                        d["pairs"] = pairs_now
                        lst = lst[:i2] + lst[i2 + 1:]
                        d[key] = lst
                        save_data(d)
                        if sb.enabled() and pairs_now:
                            try:
                                sb.save_pairs(d.get("pairs_date") or datetime.date.today().isoformat(),
                                              pairs_now)
                            except Exception as e:
                                print("[配对写入失败] %s" % e)
                        invalidate_cache()
                        self._send(200, json.dumps(
                            {"ok": True, "removed": True, "spu": spu_val},
                            ensure_ascii=False))
                        return
                    lst[i2]["same_as"] = cand_row if same else None
                    lst[i2]["mark"] = ("same" if same else "diff")
                    lst[i2]["by"] = user
                    d[key] = lst
                    save_data(d)
                    if sb.enabled() and lst[i2].get("id"):
                        try:
                            sb.update("pairs" if kind == "pair" else "news",
                                      {"id": lst[i2]["id"]},
                                      {"mark": "same" if same else "diff",
                                       "edited_by": user})
                        except Exception as e:
                            print("[Supabase 标记同步失败] %s" % e)
                self._send(200, json.dumps({"ok": True}, ensure_ascii=False))
                return

            self._send(404, json.dumps({"ok": False}, ensure_ascii=False))
        except Exception as e:
            traceback.print_exc()
            self._send(500, json.dumps({"ok": False, "msg": str(e)}, ensure_ascii=False))


class ThreadedServer(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True


def main():
    if not os.path.exists(DATA_FILE):
        print("首次运行，导入四列表…")
        print(sync_from_excel())
    print("=" * 56)
    print("  Mico 选品工作台")
    print("  监听: http://%s:%d" % (HOST, PORT))
    print("  登录: %s" % ("需要密码" if APP_PASSWORD else "无需密码"))
    print("  成员: %s" % ", ".join(USERS))
    print("=" * 56)
    srv = ThreadedServer((HOST, PORT), Handler)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")


if __name__ == "__main__":
    main()
