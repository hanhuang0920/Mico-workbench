# -*- coding: utf-8 -*-
"""
Supabase 数据层 —— 用 PostgREST REST API（无需额外依赖，仅 urllib）

配置环境变量：
  SUPABASE_URL       https://xxxxx.supabase.co
  SUPABASE_ANON_KEY  eyJhbGciOi...
"""
import os
import re
import json
import time
import contextvars
import urllib.request
import urllib.parse
import urllib.error

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_KEY = os.environ.get("SUPABASE_ANON_KEY", "")

# ---------- 多客户隔离 ----------
# 请求进入时用 set_customer() 设定当前客户，本模块所有读写自动带上该客户过滤。
CUSTOMERS = [c.strip() for c in
             os.environ.get("CUSTOMERS", "Mico,Stijn").split(",") if c.strip()]
_DEFAULT_CUSTOMER = CUSTOMERS[0] if CUSTOMERS else "Mico"
_customer_var = contextvars.ContextVar("customer", default=_DEFAULT_CUSTOMER)


def set_customer(name):
    """设定当前请求的客户（每个请求开始时调用）。"""
    if name and name in CUSTOMERS:
        _customer_var.set(name)
    else:
        _customer_var.set(_DEFAULT_CUSTOMER)
    return _customer_var.get()


def current_customer():
    return _customer_var.get()


def _cust():
    """所有按客户隔离的表都用这个值过滤。"""
    return _customer_var.get()


def enabled():
    return bool(SUPABASE_URL and SUPABASE_KEY)


class SupabaseError(Exception):
    pass


def _req(method, path, body=None, params=None, prefer=None, retries=3):
    """统一请求封装，带重试与限速。"""
    if not enabled():
        raise SupabaseError("Supabase 未配置")

    url = "%s/rest/v1/%s" % (SUPABASE_URL, path)
    if params:
        url += "?" + urllib.parse.urlencode(params)

    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": "Bearer " + SUPABASE_KEY,
        "Content-Type": "application/json",
    }
    if prefer:
        headers["Prefer"] = prefer

    data = json.dumps(body).encode("utf-8") if body is not None else None

    last = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, data=data, headers=headers, method=method)
            with urllib.request.urlopen(req, timeout=30) as r:
                raw = r.read().decode("utf-8")
                return json.loads(raw) if raw.strip() else None
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", "replace")[:400]
            last = SupabaseError("%s %s -> %d %s" % (method, path, e.code, detail))
            if e.code in (429, 500, 502, 503, 504):
                time.sleep(0.4 * (attempt + 1))
                continue
            raise last
        except Exception as e:
            last = e
            time.sleep(0.4 * (attempt + 1))
    raise SupabaseError(str(last))


# ---------- 通用 ----------
# 需要按客户隔离的表（其余表如 edit_log/op_log 保持全局）
_SCOPED_TABLES = {"products", "pairs", "news", "img_sigs"}


def _scoped(table):
    return table in _SCOPED_TABLES


def select(table, params=None):
    p = dict(params or {})
    if _scoped(table) and "customer" not in p:
        p["customer"] = "eq.%s" % _cust()
    return _req("GET", table, params=p) or []


def insert(table, rows, upsert_on=None):
    """批量插入，rows 为 list。upsert_on 指定冲突列（如 customer,url_key）实现 upsert。"""
    if not rows:
        return None
    if _scoped(table):
        cust = _cust()
        for r in rows:
            r.setdefault("customer", cust)
    params = {"on_conflict": upsert_on} if upsert_on else None
    prefer = "resolution=merge-duplicates,return=representation" if upsert_on else "return=representation"
    return _req("POST", table, body=rows, params=params, prefer=prefer)


def update(table, match, patch):
    """match: dict 过滤条件, patch: dict 更新字段"""
    params = {}
    for k, v in match.items():
        params[k] = "eq.%s" % v
    if _scoped(table):
        params["customer"] = "eq.%s" % _cust()
    return _req("PATCH", table, body=patch, params=params,
                prefer="return=representation")


def delete(table, match):
    params = {}
    for k, v in match.items():
        params[k] = "eq.%s" % v
    if _scoped(table):
        params["customer"] = "eq.%s" % _cust()
    return _req("DELETE", table, params=params)


def _req_raw(method, table, params):
    """DELETE 用（支持 gte 等操作符，不走 eq 转换）。"""
    return _req(method, table, params=params)


def count(table, params=None):
    p = dict(params or {})
    p["select"] = "id"
    p["limit"] = "1"
    # 用 HEAD + Prefer 取总数，这里简化为查询长度
    rows = select(table, {"select": "id", "limit": "100000"})
    return len(rows)


# ---------- 规范化 ----------
def norm_url(u):
    return (str(u) or "").split("#")[0].split("?")[0].rstrip("/").lower()


def split_spu(raw):
    s = re.sub(r"\s+", " ", str(raw or "")).strip()
    return [x for x in re.split(r"[,\s/;]+", s) if x]


def uniq_urls(arr):
    """链接数组去重（按规范化后的 URL，保留首次出现的顺序与原写法）。"""
    out, seen = [], set()
    for u in arr or []:
        s = str(u or "").strip()
        if not s:
            continue
        k = norm_url(s)
        if k in seen:
            continue
        seen.add(k)
        out.append(s)
    return out


# ---------- 业务：产品 ----------
def insert_products(items):
    """纯新增（不按 url_key 合并）。用于同一产品在不同站点的多个询盘号记录。"""
    rows = []
    for p in items:
        rows.append({
            "row_no": p.get("row"),
            "names": p.get("names") or [],
            "url": p.get("url") or "",
            "url_key": norm_url(p.get("url")),
            "inq": p.get("inq") or None,
            "batch": p.get("batch") or None,
            "spu": p.get("spu") or None,
            "spus": p.get("spus") or split_spu(p.get("spu")),
            "u1688": uniq_urls(p.get("u1688")),
            "img": p.get("img") or None,
        })
    for i in range(0, len(rows), 200):
        insert("products", rows[i:i + 200])
    return len(rows)


def upsert_products(items):
    """items: [{row, names, url, inq, batch, spu, spus, u1688, img}]"""
    rows = []
    for p in items:
        rows.append({
            "row_no": p.get("row"),
            "names": p.get("names") or [],
            "url": p.get("url") or "",
            "url_key": norm_url(p.get("url")),
            "inq": p.get("inq") or None,
            "batch": p.get("batch") or None,
            "spu": p.get("spu") or None,
            "spus": p.get("spus") or split_spu(p.get("spu")),
            "u1688": uniq_urls(p.get("u1688")),
            "img": p.get("img") or None,
        })
    # 分批，避免单次请求过大
    for i in range(0, len(rows), 200):
        chunk = rows[i:i + 200]
        try:
            insert("products", chunk, upsert_on="customer,url_key,inq")
        except Exception:
            try:
                insert("products", chunk, upsert_on="customer,url_key")
            except Exception:
                insert("products", chunk, upsert_on="url_key")
    return len(rows)


def get_products():
    rows = select("products", {"select": "*", "order": "row_no.asc", "limit": "10000"})
    out = []
    for r in rows:
        out.append({
            "row": r.get("row_no"),
            "id": r.get("id"),
            "names": r.get("names") or [],
            "url": r.get("url") or "",
            "inq": r.get("inq") or "",
            "batch": r.get("batch") or "",
            "spu": r.get("spu") or "",
            "spus": r.get("spus") or [],
            "multi_spu": len(r.get("spus") or []) > 1,
            "u1688": r.get("u1688") or [],
            "img": r.get("img") or "",
        })
    return out


# ---------- 业务：配对 ----------
def clear_pairs(date_str):
    delete("pairs", {"work_date": date_str})


def clear_all_pairs():
    """清空当前客户的配对（导入新订单时用，以本次订单为准）。"""
    _req("DELETE", "pairs", params={"id": "gte.0", "customer": "eq.%s" % _cust()})


def clear_all_news():
    _req("DELETE", "news", params={"id": "gte.0", "customer": "eq.%s" % _cust()})


def save_pairs(date_str, items):
    """保存配对。items 可能来自多个日期（页面上不再按日期过滤），
    所以每条优先用自己的 date 字段，避免把旧记录都改成今天。"""
    # 只清理本次涉及的日期
    dates = {date_str}
    for p in items:
        dd = p.get("date") or p.get("work_date")
        if dd:
            dates.add(str(dd))
    for dd in dates:
        clear_pairs(dd)
    rows = []
    for p in items:
        dd = p.get("date") or p.get("work_date") or date_str
        rows.append({
            "work_date": str(dd),
            "url": p.get("url"),
            "names": p.get("names") or [],
            "inq": p.get("inq") or None,
            "spu": p.get("spu") or None,
            "filled_spu": p.get("filled_spu") or None,
            "u1688": uniq_urls(p.get("u1688")),
            "img": p.get("img") or None,
            "note": p.get("note") or None,
        })
    for i in range(0, len(rows), 200):
        insert("pairs", rows[i:i + 200])
    return len(rows)


def get_pairs(date_str=None):
    params = {"select": "*", "order": "spu.asc.nullslast,id.asc", "limit": "5000"}
    if date_str:
        params["work_date"] = "eq.%s" % date_str
    rows = select("pairs", params)
    out = []
    for r in rows:
        out.append({
            "id": r.get("id"),
            "date": str(r.get("work_date") or ""),
            "work_date": str(r.get("work_date") or ""),
            "url": r.get("url") or "",
            "names": r.get("names") or [],
            "inq": r.get("inq") or "",
            "spu": r.get("spu") or "",
            "spus": split_spu(r.get("spu")),
            "filled_spu": r.get("filled_spu") or "",
            "u1688": r.get("u1688") or [],
            "img": r.get("img") or "",
            "by": r.get("edited_by") or "",
            "mark": r.get("mark"),
        })
    return out


def update_pair_spu(pair_id, spu, editor):
    patch = {"filled_spu": spu, "edited_by": editor}
    try:
        patch["spus"] = split_spu(spu)
    except Exception:
        pass
    try:
        update("pairs", {"id": pair_id}, patch)
    except Exception:
        # pairs 表可能还没有 spus 列，退回只写 filled_spu
        update("pairs", {"id": pair_id}, {"filled_spu": spu, "edited_by": editor})


# ---------- 业务：新品 ----------
def clear_news(date_str):
    delete("news", {"work_date": date_str})


def save_news(date_str, items):
    clear_news(date_str)
    rows = []
    for p in items:
        rows.append({
            "work_date": date_str,
            "url": p.get("url"),
            "name": (p.get("name") or "")[:400],
            "names": p.get("names") or [],
            "img": p.get("img") or None,
            "skus": p.get("skus") or [],
            "l1": p.get("l1") or None,
            "l2": p.get("l2") or None,
            "l3": p.get("l3") or None,
            "filled_spu": p.get("filled_spu") or None,
            "checked": p.get("checked", True),
        })
    try:
        for i in range(0, len(rows), 200):
            insert("news", rows[i:i + 200])
    except Exception as e:
        # 老库还没加 filled_spu 列时自动去掉该字段重试
        if "filled_spu" in str(e):
            for r in rows:
                r.pop("filled_spu", None)
            for i in range(0, len(rows), 200):
                insert("news", rows[i:i + 200])
        else:
            raise
    return len(rows)


def get_news(date_str=None):
    params = {"select": "*", "order": "id.asc", "limit": "5000"}
    if date_str:
        params["work_date"] = "eq.%s" % date_str
    rows = select("news", params)
    out = []
    for r in rows:
        out.append({
            "id": r.get("id"),
            "url": r.get("url") or "",
            "name": r.get("name") or "",
            "names": r.get("names") or [],
            "img": r.get("img") or "",
            "skus": r.get("skus") or [],
            "l1": r.get("l1") or "",
            "l2": r.get("l2") or "",
            "l3": r.get("l3") or "",
            "filled_spu": r.get("filled_spu") or "",
            "checked": r.get("checked", True),
        })
    return out


def update_news(item_id, patch, editor):
    patch = dict(patch)
    patch["edited_by"] = editor
    update("news", {"id": item_id}, patch)


# ---------- 业务：日志 ----------
def log_edit(editor, table_name, target_id, old_spu, new_spu, note=""):
    try:
        insert("edit_log", [{
            "editor": editor, "table_name": table_name, "target_id": target_id,
            "old_spu": old_spu, "new_spu": new_spu, "note": note,
        }])
    except Exception as e:
        print("log_edit 失败:", e)


def log_op(editor, action, detail=None):
    try:
        insert("op_log", [{"editor": editor, "action": action, "detail": detail or {}}])
    except Exception as e:
        print("log_op 失败:", e)


def get_edit_log(limit=100):
    return select("edit_log", {"select": "*", "order": "ts.desc", "limit": str(limit)}) or []


def get_stats():
    """各表数量。"""
    out = {}
    for t in ("products", "pairs", "news"):
        try:
            out[t] = count(t)
        except Exception:
            out[t] = 0
    return out


# ---------- 业务：SPU 关联的产品名称（一个 SPU = 多客户外语名/多站点） ----------
def put_spu_names(items):
    """记录 SPU→产品名/链接。items: [{spu,name,url,inq}]。表缺失时静默跳过。"""
    rows, seen = [], set()
    for x in items:
        spu = (x.get("spu") or "").strip()
        name = (x.get("name") or "").strip()[:400]
        if not spu or not name:
            continue
        key = (spu, name)
        if key in seen:
            continue
        seen.add(key)
        rows.append({"spu": spu, "name": name,
                     "url": x.get("url") or "", "inq": x.get("inq") or ""})
    if not rows:
        return 0
    try:
        for i in range(0, len(rows), 200):
            insert("spu_names", rows[i:i + 200], upsert_on="spu,name")
    except Exception as e:
        print("put_spu_names 失败（可能未建表）:", str(e)[:120])
        return 0
    return len(rows)


def get_spu_names(spu=None):
    params = {"select": "*", "limit": "5000", "order": "updated_at.desc"}
    if spu:
        params["spu"] = "eq.%s" % spu
    try:
        return select("spu_names", params) or []
    except Exception:
        return []


# ---------- 业务：图片签名缓存（提速相似款比对） ----------
def get_img_sigs():
    """全量读 img_sigs：{url_key: sig_tuple}。表不存在时返回空（优雅降级）。"""
    try:
        rows = select("img_sigs", {"select": "url_key,sig", "limit": "5000"})
    except Exception:
        return {}
    out = {}
    for r in rows or []:
        s = r.get("sig")
        if s and r.get("url_key"):
            out[r["url_key"]] = tuple(s)
    return out


_SIG_WRITE_OK = True


def put_img_sig(url_key, sig_list):
    global _SIG_WRITE_OK
    if not _SIG_WRITE_OK:
        return
    try:
        insert("img_sigs", [{"url_key": url_key, "sig": sig_list,
                             "customer": _cust()}],
               upsert_on="customer,url_key")
    except Exception as e:
        # 表不存在/权限问题时熔断，避免每张图都白白请求一次
        _SIG_WRITE_OK = False
        print("put_img_sig 失败（已停用签名回写）:", e)


def put_img_sigs_bulk(pairs_):
    """批量写入 [(url_key, sig_list)]，分批 200。返回写入数。"""
    cust = _cust()
    rows = [{"url_key": k, "sig": v, "customer": cust} for k, v in pairs_ if k and v]
    n = 0
    for i in range(0, len(rows), 200):
        try:
            insert("img_sigs", rows[i:i + 200], upsert_on="customer,url_key")
        except Exception:
            insert("img_sigs", rows[i:i + 200], upsert_on="url_key")
        n += len(rows[i:i + 200])
    return n
