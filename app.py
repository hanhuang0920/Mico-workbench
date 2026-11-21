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
from urllib.parse import parse_qs, quote, unquote

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
PER_PAGE = 48

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
            pairs = sb.get_pairs(today)
            news = sb.get_news(today)
            if not pairs and not news:
                pairs = sb.get_pairs()
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
        u1688 = [x.strip() for x in str(wi.cell(row=r, column=6).value or "").split("\n") if x.strip().startswith("http")]
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
    d = load_data()
    if not d["products"]:
        sync_from_excel()
        d = load_data()
    inq_by_url = {norm(p["url"]): p for p in d["products"]}
    prods = read_order(path)
    pairs, not_inq, with_sku = [], 0, 0
    for k, p in prods.items():
        hit = inq_by_url.get(k)
        if not hit:
            not_inq += 1
            continue
        if p["skus"]:
            with_sku += 1
            continue
        pairs.append({
            "url": p["url"],
            "names": hit["names"] + [n for n in p["names"] if n not in hit["names"]],
            "inq": hit["inq"], "spu": hit["spu"],
            "u1688": hit["u1688"], "img": hit["img"] or p["img"],
            "filled_spu": "", "note": "",
        })
    pairs.sort(key=lambda x: (0 if x["spu"] else 1, x["inq"]))
    d["pairs"] = pairs
    d["pairs_date"] = datetime.date.today().isoformat()
    if sb.enabled():
        try:
            sb.save_pairs(d["pairs_date"], pairs)
            invalidate_cache()
        except Exception as e:
            print("[Supabase 配对写入失败] %s" % e)

    # 未询盘 → 询盘需求表候选
    news = []
    for k, p in prods.items():
        if k in inq_by_url:
            continue
        nm = p["names"][0] if p["names"] else ""
        news.append({
            "url": p["url"], "name": nm, "names": p["names"],
            "img": p["img"], "skus": p["skus"],
            "l1": "", "l2": "", "l3": "", "checked": True,
        })
    d["news"] = news
    if sb.enabled():
        try:
            sb.save_news(d["pairs_date"], news)
            invalidate_cache()
        except Exception as e:
            print("[Supabase 新品写入失败] %s" % e)
    save_data(d)
    return {"ok": True, "total": len(prods), "inq": len(prods) - not_inq,
            "with_sku": with_sku, "pairs": len(pairs), "not_inq": not_inq,
            "news": len(news)}


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
        return {"ok": True, "added": 0, "total": len(quotes), "msg": "没有新产品需要导入"}

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
    return {"ok": True, "added": len(todo), "total": len(quotes),
            "imgs": len(got), "backup": bak, "total_rows": total}


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
        ws.cell(row=r, column=6, value="\n".join(p["u1688"]) if p["u1688"] else None)
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


def record_edit(pairs, editor, arr):
    """记录谁在什么时候改了什么。"""
    if sb.enabled():
        try:
            for x in arr:
                i2 = x.get("idx")
                spu = (x.get("spu") or "").strip()
                if not spu or not isinstance(i2, int) or not (0 <= i2 < len(pairs)):
                    continue
                old = pairs[i2].get("spu") or ""
                nm = (pairs[i2].get("names") or [""])[0][:80]
                sb.log_edit(editor, "pairs", pairs[i2].get("id"),
                            old, spu, nm)
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


def apply_spu_to_master(user_name=""):
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
            # 配对表填写的 SPU 也写回 Supabase 的 products 表
            prods_now = []
            for p in filled:
                prods_now.append({
                    "row": None,
                    "names": p.get("names") or [],
                    "url": p.get("url") or "",
                    "inq": p.get("inq") or "",
                    "spu": p["filled_spu"].strip(),
                    "spus": sb.split_spu(p["filled_spu"]),
                    "u1688": p.get("u1688") or [],
                    "img": p.get("img") or "",
                })
            if prods_now:
                sb.upsert_products(prods_now)
            # 同步 pairs 表里的人工填写值
            for p in filled:
                if p.get("id"):
                    try:
                        sb.update_pair_spu(p["id"], p["filled_spu"].strip(), user_name)
                    except Exception:
                        pass
        except Exception as e:
            print("[Supabase SPU 同步失败] %s" % e)
    invalidate_cache()
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


def product_card(p):
    name = p["names"][0] if p["names"] else ""
    extra = len(p["names"]) - 1
    if p.get("multi_spu"):
        spu_html = "".join('<b class="spu multi">%s</b>' % esc(s) for s in p["spus"])
        spu_html += '<b class="tag">多</b>'
    elif p["spu"]:
        spu_html = '<b class="spu">%s</b>' % esc(p["spu"])
    else:
        spu_html = '<b class="spu none">无 SPU</b>'
    img = ('<img src="%s" loading="lazy" decoding="async">' % esc(p["img"])) if p["img"] else '<span class="nopic">无图</span>'
    link = ('<a href="%s" target="_blank" rel="noopener">站点</a>' % esc(p["url"])) if p["url"].startswith("http") else ""
    l1688 = ('<a href="%s" target="_blank" rel="noopener">1688</a>' % esc(p["u1688"][0])) if p["u1688"] else ""
    return (
        '<div class="pc" data-n="%s" data-s="%s" data-i="%s" data-m="%s" data-p="%s" data-q="%s">'
        '<div class="pi">%s</div><div class="pb">'
        '<div class="pn" title="%s">%s%s</div>'
        '<div class="ps">%s</div>'
        '<div class="pf"><span class="iq">%s</span></div>'
        '<div class="pl">%s %s</div></div></div>' % (
            esc(name.lower()), esc(p["spu"].lower()), esc(p["inq"].lower()),
            "1" if p.get("multi_spu") else "0",
            "0" if p["spu"] else "1",
            "1" if not p["inq"] else "0",
            img, esc(" ／ ".join(p["names"])), esc(name[:90]),
            ('<i>+%d</i>' % extra) if extra > 0 else "",
            spu_html, esc(p["inq"][-14:] if p["inq"] else "—"), link, l1688))


def pair_card(i, p):
    img = ('<img src="%s" loading="lazy">' % esc(p["img"])) if p["img"] else '<span class="nopic">无图</span>'
    cur = p.get("filled_spu") or p.get("spu") or ""
    return (
        '<div class="paircard"><div class="pi">%s</div><div class="pb">'
        '<div class="pn">%s</div>'
        '<div class="ps"><span class="iq">%s</span></div>'
        '<div class="inputrow"><label>SPU</label>'
        '<input class="spin" data-idx="%d" value="%s" placeholder="填写 SPU"></div>'
        '<div style="display:flex;gap:6px;margin-top:4px">'
        '<button class="simbtn%s" onclick="sim(%d,\'pair\')">相似款比对</button>'
        '</div>'
        '</div></div>' % (
            img, esc(p["names"][0][:84] if p["names"] else ""),
            esc((p.get("inq") or "")[-14:]), i, esc(cur),
            " done" if p.get("mark") == "same" else "", i))


def news_card(i, p):
    img = ('<img src="%s" loading="lazy">' % esc(p["img"])) if p.get("img") else '<span class="nopic">无图</span>'
    checked = "checked" if p.get("checked", True) else ""
    return (
        '<div class="newscard" data-idx="%d">'
        '<label class="chk"><input type="checkbox" data-idx="%d" %s></label>'
        '<div class="pi">%s</div><div class="pb">'
        '<div class="pn" title="%s">%s</div>'
        '<div class="cat3">'
        '<input class="ci" data-idx="%d" data-k="l1" placeholder="一级" value="%s">'
        '<input class="ci" data-idx="%d" data-k="l2" placeholder="二级" value="%s">'
        '<input class="ci" data-idx="%d" data-k="l3" placeholder="三级" value="%s">'
        '</div>'
        '<div class="pl"><a href="%s" target="_blank" rel="noopener">站点</a>'
        '<button class="simbtn%s" onclick="sim(%d,\'news\')">相似款</button></div>'
        '</div></div>' % (
            i, i, checked, img,
            esc(p.get("name", "")), esc(p.get("name", "")[:80]),
            i, esc(p.get("l1", "")), i, esc(p.get("l2", "")), i, esc(p.get("l3", "")),
            esc(p.get("url", "")),
            " done" if p.get("mark") == "same" else "", i))


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


def page_index():
    d = load_data()
    prods = d.get("products", [])
    pairs = d.get("pairs", [])
    n_spu = sum(1 for p in prods if p["spu"])
    n_multi = sum(1 for p in prods if p.get("multi_spu"))
    n_noiq = sum(1 for p in prods if not p["inq"])

    cards = "".join(product_card(p) for p in prods[:PER_PAGE])
    pcards = "".join(pair_card(i, p) for i, p in enumerate(pairs))
    no_pairs = '<div class="empty">还没有配对数据<br><span>请在 ① 上传今日订单文件</span></div>' if not pairs else ""

    news = d.get("news", [])
    ncards = "".join(news_card(i, p) for i, p in enumerate(news))
    no_news = '<div class="empty">还没有未询盘产品<br><span>请在 ① 上传今日订单文件</span></div>' if not news else ""

    return PAGE_TPL \
        .replace("{{USER}}", esc(CURRENT_USER.get("name", ""))) \
        .replace("{{TOTAL}}", str(len(prods))) \
        .replace("{{NSPU}}", str(n_spu)) \
        .replace("{{NMULTI}}", str(n_multi)) \
        .replace("{{NNOIQ}}", str(n_noiq)) \
        .replace("{{NPAIRS}}", str(len(pairs))) \
        .replace("{{PER}}", str(PER_PAGE)) \
        .replace("{{CARDS}}", cards) \
        .replace("{{PCARDS}}", pcards) \
        .replace("{{EMPTY}}", no_pairs) \
        .replace("{{NCARDS}}", ncards) \
        .replace("{{NONEWS}}", no_news) \
        .replace("{{NNEWS}}", str(len(news)))


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
.spin{flex:1;padding:6px 9px;border:1px solid #e2e8f0;border-radius:6px;
 font-family:ui-monospace,Consolas,monospace;font-size:11.5px;outline:none;min-width:0}
.spin:focus{border-color:#4f46e5}
.spin.filled{border-color:#10b981;background:#f0fdf4}
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
 outline:none;font-family:inherit;width:100%}
.ci:focus{border-color:#4f46e5;background:#fafaff}
.ci.ok{border-color:#10b981;background:#f0fdf4}
.simbtn{padding:4px 10px;border-radius:6px;border:1px solid #c7d2fe;background:#eef2ff;
 color:#4338ca;font-size:10.5px;font-weight:600;cursor:pointer;font-family:inherit}
.simbtn:hover{background:#e0e7ff}
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
  <div class="more" id="more0"><button onclick="loadMore()">加载更多</button><span id="tip0"></span></div>
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
      <button class="btn" onclick="saveSpu()">保存 SPU 并写回四列表</button>
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
    if(d.added===0){toast('没有新产品需要导入（询盘文件共 '+d.total+' 款）');return}
    toast('已导入 '+d.added+' 款新品（图片 '+d.imgs+' 张），已备份');
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
    var l3=c.querySelector('input[data-k=l3]');
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
  }).then(function(r){return r.json()})
  .then(function(d){
    if(!d.ok){toast('失败：'+(d.msg||''));return}
    toast('已导出 '+d.file+'（'+d.count+' 款）');
  }).catch(function(e){toast('出错：'+e)});
}
var SIMKIND='pair', SIMIDX=-1;
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
    h+='<div class="cn" style="font-weight:700">'+esc2(t.names[0]||'')+'</div></div></div>';
    if(!cs.length){
      h+='<div class="simempty">已询盘库中<b>没有找到相似产品</b><br>可以放心当作新品处理</div>';
    }else{
      h+='<div style="font-size:12.5px;color:#64748b;margin-bottom:6px">找到 '
         +cs.length+' 个相似款，请确认是否为同款：</div>';
      cs.forEach(function(c,k){
        h+='<div class="cmp">';
        h+='<div class="cmpimg">'+(c.img?'<img src="'+esc2(c.img)+'">':'<span class="nopic">无图</span>')+'</div>';
        h+='<div class="cmpinfo">';
        h+='<div class="cn">'+esc2(c.name)+'</div>';
        h+='<div class="cs">'+spuHtml(c)+'</div>';
        h+='<div style="font-size:11px;color:#94a3b8">'
           +(c.dist!==null?('图距 '+c.dist+' · '):'')
           +'标题相似 '+c.nsim+'</div>';
        h+='<div class="pl"><a href="'+esc2(c.url)+'" target="_blank">站点</a>'
           +'<span class="iq">'+esc2((c.inq||'').slice(-14))+'</span></div>';
        h+='</div>';
        h+='<div class="cmpact">';
        h+='<span class="level '+c.level+'">'+esc2(c.level_txt)+'</span>';
        h+='<button class="byes" onclick="markSame('+idx+',\''+kind+'\','+c.row+')">是同一款</button>';
        h+='<button class="bno" onclick="markDiff('+idx+',\''+kind+'\')">不是同款</button>';
        h+='</div></div>';
      });
    }
    b.innerHTML=h;
  }).catch(function(e){b.innerHTML='<div class="simempty">出错：'+esc2(e)+'</div>'});
}
function markSame(idx,kind,row){
  fetch('/api/mark-same',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({idx:idx,kind:kind,same:true,cand_row:row})})
  .then(function(r){return r.json()}).then(function(d){
    toast(d.ok?'已标记为同款，可据此剔除':'保存失败');
    closeModal();
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
                self._send(200, json.dumps(apply_spu_to_master(user), ensure_ascii=False))
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

            if self.path == "/api/export":
                self._send(200, json.dumps(export_pairs(), ensure_ascii=False))
                return

            if self.path == "/api/export-demand":
                self._send(200, json.dumps(export_demand(), ensure_ascii=False))
                return

            if self.path == "/api/similar":
                payload = json.loads(raw.decode("utf-8"))
                kind = payload.get("kind", "pair")
                i2 = payload.get("idx")
                d = load_data()
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
                                              exclude_spu=_sp)
                out = []
                for c in cands:
                    c["level_txt"] = similar.LEVEL_TXT.get(c["level"], "")
                    out.append(c)
                self._send(200, json.dumps(
                    {"ok": True, "target": {"names": names, "img": t.get("img", "")},
                     "candidates": out}, ensure_ascii=False))
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
