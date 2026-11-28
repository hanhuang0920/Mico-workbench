# -*- coding: utf-8 -*-
"""相似款匹配引擎：图片哈希 + 标题相似度，返回候选及 SPU。"""
import sys, io, os, re, json, threading, urllib.request
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from difflib import SequenceMatcher

BASE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(BASE, "_imgcache")


def norm(u):
    return (str(u) or "").split("#")[0].split("?")[0].rstrip("/").lower()


def ensure_img(url, timeout=15):
    """下载图片到缓存，返回本地路径；失败返回 None。"""
    if not url or not url.startswith("http"):
        return None
    os.makedirs(CACHE, exist_ok=True)
    key = re.sub(r"[^A-Za-z0-9]", "_", url)[-90:] + ".img"
    fp = os.path.join(CACHE, key)
    if os.path.exists(fp) and os.path.getsize(fp) > 500:
        return fp
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            d = r.read()
        if len(d) < 500:
            return None
        open(fp, "wb").write(d)
        return fp
    except Exception:
        return None


try:
    from PIL import Image as PILImage
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


def _ah(p, s=16):
    im = PILImage.open(p).convert("L").resize((s, s), PILImage.LANCZOS)
    px = list(im.getdata())
    a = sum(px) / len(px)
    b = 0
    for v in px:
        b = (b << 1) | (1 if v > a else 0)
    return b


def _dh(p, s=16):
    im = PILImage.open(p).convert("L").resize((s + 1, s), PILImage.LANCZOS)
    px = list(im.getdata())
    w = s + 1
    b = 0
    for y in range(s):
        for x in range(s):
            b = (b << 1) | (1 if px[y * w + x] > px[y * w + x + 1] else 0)
    return b


SIG_CACHE_FILE = os.path.join(BASE, "_sigs.json")
_sig_mem = None


def _load_sigs():
    global _sig_mem
    if _sig_mem is None:
        if os.path.exists(SIG_CACHE_FILE):
            try:
                raw = json.load(open(SIG_CACHE_FILE, encoding="utf-8"))
                _sig_mem = {k: tuple(v) for k, v in raw.items()}
            except Exception:
                _sig_mem = {}
        else:
            _sig_mem = {}
    return _sig_mem


def _save_sigs():
    if _sig_mem is not None:
        try:
            json.dump({k: list(v) for k, v in _sig_mem.items()},
                      open(SIG_CACHE_FILE, "w", encoding="utf-8"))
        except Exception:
            pass


def _crop_subject(path, pad=0.06):
    """
    裁剪出产品主体：去掉与四角颜色接近的背景区域。
    返回 PIL Image 或 None。
    """
    try:
        from PIL import ImageChops
    except ImportError:
        return None
    try:
        im = PILImage.open(path).convert("RGB")
        # 统一缩到 200x200 便于处理
        im = im.resize((200, 200), PILImage.LANCZOS)
        w, h = im.size
        # 采样四角确定背景色
        corners = [im.getpixel((2, 2)), im.getpixel((w - 3, 2)),
                   im.getpixel((2, h - 3)), im.getpixel((w - 3, h - 3))]
        # 只取最暗/最亮的少数作背景（背景通常占比大且一致）
        bg = corners[0]
        # 计算与背景的差异图
        bgimg = PILImage.new("RGB", im.size, bg)
        diff = ImageChops.difference(im, bgimg).convert("L")
        # 二值化：显著差异 = 主体
        thr = max(18, int(sum(diff.getdata()) / (w * h) * 2.2))
        mask = diff.point(lambda v: 255 if v > thr else 0)
        bbox = mask.getbbox()
        if not bbox:
            return im
        x0, y0, x1, y1 = bbox
        dw, dh = x1 - x0, y1 - y0
        # 过小的框不可信
        if dw < w * 0.15 or dh < h * 0.15:
            return im
        pd = int(max(dw, dh) * pad)
        box = (max(0, x0 - pd), max(0, y0 - pd),
               min(w, x1 + pd), min(h, y1 + pd))
        return im.crop(box)
    except Exception:
        return None


def _shape_hash(path, s=16):
    """
    形状哈希：裁剪主体 → 方形归一化 → 边缘二值化 → 逐格比较。
    对尺寸、裁剪比例、颜色都不敏感（只看轮廓结构）。
    """
    try:
        from PIL import ImageFilter
        from PIL import ImageOps
    except ImportError:
        return None
    try:
        im = _crop_subject(path)
        if im is None:
            im = PILImage.open(path).convert("RGB")
        # 转灰度并做边缘增强
        g = im.convert("L")
        # 方形补边（保持比例）
        g = ImageOps.pad(g, (s, s), color=128, method=PILImage.LANCZOS)
        g = g.filter(ImageFilter.FIND_EDGES)
        px = list(g.getdata())
        avg = sum(px) / len(px) * 1.05
        b = 0
        for v in px:
            b = (b << 1) | (1 if v > avg else 0)
        return b
    except Exception:
        return None


def _edge_hash(path, s=16):
    """结构哈希：先转灰度再做边缘检测，对颜色不敏感，关注形状/轮廓。"""
    try:
        from PIL import ImageFilter
    except ImportError:
        return None
    try:
        im = PILImage.open(path).convert("L").resize((s, s), PILImage.LANCZOS)
        im = im.filter(ImageFilter.FIND_EDGES)
        px = list(im.getdata())
        avg = sum(px) / len(px)
        b = 0
        for v in px:
            b = (b << 1) | (1 if v > avg else 0)
        return b
    except Exception:
        return None


def _sat_hash(path, s=16):
    """饱和度哈希：HSV 的 S 通道，弱化亮度/色相，对"同款不同色"友好。"""
    try:
        im = PILImage.open(path).convert("HSV").resize((s, s), PILImage.LANCZOS)
        px = list(im.split()[1].getdata())
        avg = sum(px) / len(px)
        b = 0
        for v in px:
            b = (b << 1) | (1 if v > avg else 0)
        return b
    except Exception:
        return None


def _sig32(path, s=32):
    """32x32 精细签名（灰度均值哈希 + 梯度哈希），对构图比例比16x16鲁棒。"""
    try:
        import numpy as np
    except ImportError:
        return None
    try:
        im = PILImage.open(path).convert("RGB")
        g = np.asarray(im.convert("L").resize((s, s), PILImage.LANCZOS), dtype=np.float32)
        m = g.mean()
        b = 0
        for v in g.ravel():
            b = (b << 1) | (1 if v > m else 0)
        g2 = np.asarray(im.convert("L").resize((s + 1, s), PILImage.LANCZOS), dtype=np.float32)
        d = g2[:, 1:] > g2[:, :-1]
        b2 = 0
        for v in d.ravel():
            b2 = (b2 << 1) | int(v)
        return (b, b2)
    except Exception:
        return None


_white_cache = {}


def whiteness(path):
    """从签名缓存读取白底度（避免重复打开图片）。"""
    sig = img_sig(path) if path else None
    if sig and len(sig) >= 9:
        return sig[7] / 1000.0
    return _whiteness_raw(path)


def _whiteness_raw(path):
    """返回白底程度 0~1：白底图色相信息无意义，需降低色相权重。（带缓存）"""
    if not path:
        return 0.0
    try:
        key = "%s|%d" % (path, int(os.path.getmtime(path)))
    except Exception:
        return 0.0
    if key in _white_cache:
        return _white_cache[key]
    try:
        import numpy as np
    except ImportError:
        return 0.0
    try:
        im = PILImage.open(path).convert("RGB").resize((48, 48), PILImage.LANCZOS)
        a = np.asarray(im, dtype=np.float32)
        # 统计接近白色的像素比例
        white = np.all(a > 225, axis=2).mean()
        bright = (a.mean(axis=2) > 235).mean()
        v = float(max(white, bright))
        _white_cache[key] = v
        return v
    except Exception:
        return 0.0


def hue_hist(path, bins=12):
    """从签名缓存读取色相直方图。"""
    sig = img_sig(path) if path else None
    if sig and len(sig) >= 9:
        try:
            import numpy as np
            return np.array(sig[8], dtype=np.float32)
        except Exception:
            return None
    return _hue_hist_raw(path, bins)


def _hue_hist_raw(path, bins=12):
    """色相直方图（对"同款同图不同色"场景的辅助证据）。"""
    try:
        import numpy as np
    except ImportError:
        return None
    try:
        im = PILImage.open(path).convert("RGB").resize((64, 64), PILImage.LANCZOS)
        a = np.asarray(im.convert("HSV"), dtype=np.float32)
        hh, _ = np.histogram(a[:, :, 0], bins=bins, range=(0, 256))
        hh = hh.astype(np.float32)
        s = hh.sum()
        return hh / s if s > 0 else hh
    except Exception:
        return None


def cosv(a, b):
    try:
        import numpy as np
    except ImportError:
        return 0.0
    if a is None or b is None:
        return 0.0
    na, nb = float(np.linalg.norm(a)), float(np.linalg.norm(b))
    return 0.0 if na == 0 or nb == 0 else float(np.dot(a, b) / (na * nb))


def d32(a, b):
    if not a or not b:
        return 9999
    return bin(a[0] ^ b[0]).count("1") + bin(a[1] ^ b[1]).count("1")


def _all_from_im(im):
    """从已打开的 Image 一次性算出全部特征（避免重复解码）。"""
    from PIL import ImageFilter, ImageOps
    import numpy as np

    im_rgb = im.convert("RGB")

    # ---- aHash 16x16 ----
    g16 = im_rgb.convert("L").resize((16, 16), PILImage.LANCZOS)
    px = list(g16.getdata())
    avg = sum(px) / len(px)
    ahash = 0
    for v in px:
        ahash = (ahash << 1) | (1 if v > avg else 0)

    # ---- dHash 16x16 ----
    g17 = im_rgb.convert("L").resize((17, 16), PILImage.LANCZOS)
    p17 = list(g17.getdata())
    dhash = 0
    for y in range(16):
        for x in range(16):
            dhash = (dhash << 1) | (1 if p17[y * 17 + x] > p17[y * 17 + x + 1] else 0)

    # ---- 边缘哈希 ----
    e16 = im_rgb.convert("L").resize((16, 16), PILImage.LANCZOS).filter(ImageFilter.FIND_EDGES)
    epx = list(e16.getdata())
    eavg = sum(epx) / len(epx)
    edge = 0
    for v in epx:
        edge = (edge << 1) | (1 if v > eavg else 0)

    # ---- 饱和度哈希 ----
    s16 = im_rgb.convert("HSV").resize((16, 16), PILImage.LANCZOS).split()[1]
    spx = list(s16.getdata())
    savg = sum(spx) / len(spx)
    sat = 0
    for v in spx:
        sat = (sat << 1) | (1 if v > savg else 0)

    # ---- 形状哈希（主体裁剪 + 方形归一 + 边缘）----
    shape = 0
    try:
        tmp = os.path.join(BASE_CACHE, "_shape_tmp.png")
        im_rgb.save(tmp)
        crop = _crop_subject(tmp)
        if crop is None:
            crop = im_rgb
        g = crop.convert("L")
        g = ImageOps.pad(g, (16, 16), color=128, method=PILImage.LANCZOS)
        g = g.filter(ImageFilter.FIND_EDGES)
        spx2 = list(g.getdata())
        savg2 = sum(spx2) / len(spx2) * 1.05
        for v in spx2:
            shape = (shape << 1) | (1 if v > savg2 else 0)
    except Exception:
        shape = 0

    # ---- 32x32 精细 ----
    g32 = np.asarray(im_rgb.convert("L").resize((32, 32), PILImage.LANCZOS), dtype=np.float32)
    m32 = g32.mean()
    f1 = 0
    for v in g32.ravel():
        f1 = (f1 << 1) | (1 if v > m32 else 0)
    g33 = np.asarray(im_rgb.convert("L").resize((33, 32), PILImage.LANCZOS), dtype=np.float32)
    d33 = g33[:, 1:] > g33[:, :-1]
    f2 = 0
    for v in d33.ravel():
        f2 = (f2 << 1) | int(v)

    # ---- 白底度 ----
    a48 = np.asarray(im_rgb.resize((48, 48), PILImage.LANCZOS), dtype=np.float32)
    white = float(max(np.all(a48 > 225, axis=2).mean(),
                      (a48.mean(axis=2) > 235).mean()))

    # ---- 色相直方图（12 桶）----
    hsv48 = np.asarray(im_rgb.resize((64, 64), PILImage.LANCZOS).convert("HSV"),
                       dtype=np.float32)[:, :, 0]
    hh, _ = np.histogram(hsv48, bins=12, range=(0, 256))
    hsum = hh.sum()
    hue = tuple((hh / hsum).tolist() if hsum else hh.tolist())

    return (ahash, dhash, edge, sat, shape, f1, f2, int(white * 1000), hue)


BASE_CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_imgcache")
_white_cache2 = {}


def _all_from_path(path):
    im = PILImage.open(path)
    return _all_from_im(im)


def img_sig(path):
    """带缓存的图片特征（一次解码算全部）。"""
    if not HAS_PIL or not path:
        return None
    try:
        key = "%s|%d" % (path, int(os.path.getmtime(path)))
    except Exception:
        return None
    cache = _load_sigs()
    if key in cache and len(cache[key]) >= 9:
        return cache[key]
    try:
        sig = _all_from_path(path)
    except Exception:
        return None
    cache[key] = sig
    return sig


# ===== URL 级签名缓存（存 Supabase img_sigs 表，云端重启不用重新下载图片）=====
# 按客户分开缓存，避免 A 客户用到 B 客户的签名
_URL_SIGS = {}
_sig_lock = threading.Lock()
_sig_inflight = {}
_SIGS_CHANGED = False


def _load_url_sigs():
    try:
        import supabase_store as _sb
        cust = _sb.current_customer() if hasattr(_sb, "current_customer") else "Mico"
    except Exception:
        return {}
    if cust not in _URL_SIGS:
        _URL_SIGS[cust] = {}
        try:
            import supabase_store as _sb
            if _sb.enabled():
                _URL_SIGS[cust] = _sb.get_img_sigs()
        except Exception as e:
            print("[img_sigs 读取失败，退回在线下载] %s" % e)
    return _URL_SIGS[cust]


def clear_sig_cache():
    """切换客户时清掉内存签名缓存。"""
    _URL_SIGS.clear()


def sig_for_url(url):
    """优先用 Supabase 签名缓存（不下载图片）；缺失才下载计算并写回。
    返回 (sig, path_or_None)：path 为 None 表示无需本地文件。"""
    global _SIGS_CHANGED
    u = norm(url)
    if not u or not str(url).startswith("http"):
        return None, None
    sigs = _load_url_sigs()
    cached = sigs.get(u)
    if cached is not None and len(cached) >= 9:
        return tuple(cached), None
    # 同一个 URL 正在被别的线程算，等它算完直接用结果（避免重复下载）
    with _sig_lock:
        memo = _sig_inflight.get(u)
        if memo is not None:
            ev, box = memo
        else:
            ev, box = threading.Event(), {}
            _sig_inflight[u] = (ev, box)
            memo = None
    if memo is not None:
        ev.wait(timeout=25)
        v = box.get("sig")
        if v:
            return v, None
        # 对方失败了，自己算
    p = ensure_img(url)
    sig = img_sig(p) if p else None
    if sig is not None:
        sigs[u] = sig
        _SIGS_CHANGED = True
        try:
            import supabase_store as _sb
            if _sb.enabled():
                _sb.put_img_sig(u, [x if not isinstance(x, tuple) else list(x) for x in sig])
        except Exception:
            pass
    with _sig_lock:
        ent = _sig_inflight.pop(u, None)
        if ent:
            ent[1]["sig"] = tuple(sig) if sig is not None else None
            ent[0].set()
    return sig, p


def _hue_of(sig):
    if sig and len(sig) >= 9 and sig[8]:
        try:
            import numpy as np
        except ImportError:
            return None
        return np.array(sig[8], dtype=np.float32)
    return None


def _white_of(sig):
    if sig and len(sig) >= 8:
        return sig[7] / 1000.0
    return 0.0


def hdist(a, b):
    """亮度哈希距离（aHash + dHash），对颜色敏感。"""
    if not a or not b:
        return 9999
    return bin(a[0] ^ b[0]).count("1") + bin(a[1] ^ b[1]).count("1")


def sdist(a, b):
    """结构距离（边缘 + 饱和度），对颜色不敏感。"""
    if not a or not b or len(a) < 4 or len(b) < 4:
        return 9999
    e = bin(a[2] ^ b[2]).count("1")
    s = bin(a[3] ^ b[3]).count("1")
    return e + s


def fine_dist(a, b):
    """32x32 精细距离（上限 2048），对裁剪/尺寸差异更宽容。"""
    if not a or not b or len(a) < 7 or len(b) < 7:
        return 9999
    return bin(a[5] ^ b[5]).count("1") + bin(a[6] ^ b[6]).count("1")


def shapdist(a, b):
    """形状距离：仅比对裁剪归一化后的轮廓，对尺寸/裁剪/颜色都鲁棒。
    shape 为 0 表示该图计算失败（不是真的相同），必须按无效值处理，否则两张失败图会被误判同款。"""
    if not a or not b or len(a) < 5 or len(b) < 5:
        return 9999
    if not a[4] or not b[4]:
        return 9999
    return bin(a[4] ^ b[4]).count("1")


STOP = set("mit und fur für der die das dem den des ein eine oder aus von zum zur "
           "na do za z ze w oraz i o od po przy dla przez z with and for the of in on "
           "to a an set pcs cm mm led usb new".split())


def toks(s):
    return {w for w in re.split(r"[^a-zA-Zäöüßąćęłńóśźż0-9]+", str(s).lower())
            if len(w) > 2 and w not in STOP}


def nsim(a, b):
    ta, tb = toks(a), toks(b)
    if not ta or not tb:
        return 0.0
    jac = len(ta & tb) / len(ta | tb)
    seq = SequenceMatcher(None, str(a).lower(), str(b).lower()).ratio() * 0.85
    return max(jac, seq)


# ===== 跨语言产品品类词表（德/英/波兰语词根 → 中文品名） =====
# 顺序 = 优先级：具体品类在前，宽泛品类在后；cn_of() 取第一个命中。
CATEGORY_WORDS = {
    # ---- 看图核对新增（2026-11-21）----
    "坐垫": ["sitzkissen", "seat cushion", "坐垫", "办公椅垫"],
    "桌垫": ["podkładka na biurko", "desk mat", "桌垫", "书桌垫"],
    "老花镜": ["okulary do czytania", "reading glasses", "老花镜"],
    "户外淋浴袋": ["torba prysznicowa", "camping shower", "淋浴袋"],
    "超声波清洗机": ["ultraschallreinigungs", "ultrasonic cleaner", "超声波清洗"],
    "电热袜": ["podgrzewane skarpetki", "beheizbare socken", "heated socks", "电热袜"],
    "烟灰缸": ["popielniczka", "ashtray", "烟灰缸"],
    "矫形鞋": ["buty ortopedyczne", "orthopedic shoes", "矫形鞋"],
    "悬浮足球": ["hoverball", "air soccer", "悬浮足球"],
    "精油加热器": ["massageölwärmer", "oil warmer", "精油加热"],
    "宠物体温计": ["termometr dla zwierząt", "für haustiere", "dla zwierząt", "宠物体温"],
    "耳环": ["earring", "stud earrings", "kolczyk", "耳环", "耳钉"],
    "量角器": ["winkelmesser", "angle finder", "neigungsmesser", "量角器", "坡度仪"],
    # ---- 宠物穿戴（最先判，含 dla psa / Hund 等词，防止被人类服装抢走）----
    "宠物服装": ["dla psa", "dla kota", "hundemantel", "hundeparka", "hundejacke", "ubranka",
             "pet clothing", "dog sweater", "dog jacket", "kurtka dla psa", "parka dla psa",
             "宠物服", "狗衣", "狗雨衣"],
    # ---- 服装（具体在前）----
    "羊绒开衫": ["kaschmir-cardigan", "cardigan", "kardigan"],
    "羊绒衫": ["kaschmir pullover", "kaschmirpullover", "kaszmir pullover", "cashmere sweater",
             "strickpullover", "merinos", "羊绒衫"],
    "毛衣针织衫": ["sweter", "pullover", "sweater", "针织衫"],
    "卫衣帽衫": ["bluza", "hoodie", "sweatshirt", "z kapturem"],
    "衬衫": ["koszula", "henley", "shirt", "hemd", "衬衫"],
    "连衣裙": ["sukienka", "winterkleid", "dress", "kleid", "连衣裙"],
    "半身裙": ["spódnica", "skirt"],
    "连体裤": ["jumpsuit", "overal"],
    "T恤": ["t-shirt", "tshirt", "koszulka", "tee "],
    "马甲背心": ["kamizelka", "sleeveless vest", "weste", "vest"],
    "派克大衣": ["parka", "派克"],
    "雨衣雨披": ["regenjacke", "raincoat", "przeciwdeszczowa", "poncho", "雨衣"],
    "皮衣": ["skórzana kurtka", "kurtka damska skórzana", "kurtka skórzana",
          "lederjacke", "leather jacket"],
    "冲锋衣外套": ["winddichte outdoorjacke", "outdoorjacke", "jacke",
               "kurtka", "jack", "外套", "夹克", "steppmantel", "mantel", "coat"],
    "抓绒衣": ["fleecejacke", "fleece jacket"],
    "斗篷毯": ["kapuzendecke", "poncho", "cape"],
    "披肩": ["shawl wrap", "shawl", "pashmina", "chusta", "披肩"],
    "打底裤": ["legginsy", "leggings", "leggings", "winterleggings", "fleeceleggings", "打底裤"],
    "牛仔裤": ["jeans"],
    "休闲裤": ["spodnie", "pants", "baggy", "裤"],
    "内裤内衣": ["biustonosz", "bra ", "boxer", "underwear", "内衣"],
    "保暖裤护腰": ["nierenwärmer", "niere", "护腰"],
    # ---- 穿戴配件 ----
    "围巾脖套": ["schlauchschal", "szalik", "winterschal", "komin", "scarf", "围巾"],
    "巴拉克拉法帽": ["kominiarka", "sturmhaube", "balaclava"],
    "帽子": ["beanie", "mütze", "mutze", "czapka", "kapelusz", "hood",
           "hat", "帽子", "bucket", "fedora"],
    "耳罩": ["ohrenschützer", "earmuff", "耳罩"],
    "手套": ["rękawiczki", "rekawiczki", "rękawice", "handschuhe", "gloves", "手套"],
    "袜子": ["socken", "skarpetki", "socks", "腿套", "beinstulpen", "袜子"],
    "连裤袜丝袜": ["strumpfhose", "pantyhose", "tights"],
    "腰带": ["gürtel", "belt", "gürtel ohne", "腰带"],
    "领带套装": ["krawat", "tie set", "领结"],
    "墨镜太阳镜": ["sunglasses", "sonnenbrille", "okulary", "goggles", "nachtfahrbrille", "护目镜"],
    "泳镜": ["swimming goggles", "best swimming"],
    # ---- 鞋 ----
    "拖鞋": ["hausschuhe", "kapcie", "slippers", "schlappen", "拖鞋"],
    "雪地靴冬鞋": ["stiefel", "botki", "buty zimowe", "snow boots", "winter shoes", "靴"],
    "凉鞋": ["sandals", "sandały"],
    "运动鞋": ["sneakers", "tennisówki", "running shoes", "运动鞋"],
    "鞋垫": ["wkładki do butów", "wkładki", "insole", "鞋垫"],
    "宠物鞋": ["buty dla psa", "dog boots", "狗鞋"],
    # ---- 包 ----
    "背包书包": ["rucksack", "backpack", "plecak", "背包", "书包"],
    "斜挎包": ["listonoszka", "shoulder bag", "torba na ramię", "saszetka", "crossbody", "挎包"],
    "手提包": ["handbag", "torebka", "torba damska", "手提包"],
    "化妆包": ["kosmetyczka", "makeup bag", "化妆包"],
    "工具包": ["torba narzędziowa", "tool bag", "工具包"],
    "收纳包": ["car seat organiser", "organiser", "organizer bag", "收纳包"],
    "防火文件袋": ["fireproof", "document bag"],
    # ---- 家居床品 ----
    "婴儿枕": ["baby pillow", "reflux pillow", "noworodka", "niemowląt", "niemowl",
             "pozycjonująca", "poduszka dziecięca", "miękka poduszka dziecięca",
             "婴儿枕", "反流枕"],
    "枕头": ["poduszka do snu", "poduszka ortopedyczna", "poduszka podr", "neck pillow",
        "na szyję", "旅行枕", "U型枕", "headrest", "pillow", "kissen", "poduszka", "枕头"],
    "抱枕靠垫": ["throw pillow", "dekorative poduszka", "抱枕", "靠垫"],
    "床垫保护罩": ["mattress", "matraze", "床罩"],
    "毯子": ["koc", "blanket", "decke", "毯"],
    "四件套": ["pościel", "bedding", "被套"],
    "地毯地垫": ["teppich", "dywan", "地毯", "地垫"],
    "浴室防滑垫": ["duschmatte", "shower mat", "mata prysznicowa", "浴室垫"],
    "冰垫宠物垫": ["cooling mat", "冰垫", "宠物冰垫"],
    "加热垫": ["podkładka grzewcza", "poduszka grzewcza", "podgrzewane podkładki", "heizkissen", "electric blanket",
            "加热垫", "grzewcza na ból", "koc grzewczy", "电热垫", "电热毯"],
    "窗帘": ["vorhang", "zasłona", "curtain", "窗帘"],
    "毛巾": ["handtuch", "ręcznik", "towel", "毛巾"],
    "浴刷": ["szczotka do mycia", "bath brush", "洗澡刷"],
    "马桶刷": ["toilet brush", "马桶刷"],
    "挂钩": ["clothes hanger", "wall hanger", "coat rack", "haken", "haczyk", "wieszak", "hook", "挂钩", "挂衣"],
    "收纳盒": ["aufbewahrung", "pudełko", "storage box", "aufbewahrungs", "收纳"],
    "书架置物架": ["regal", "półka", "shelf", "书架", "置物架"],
    "果盘": ["miska na owoce", "fruit tray", "果盘"],
    "核桃夹": ["nussknacker", "nutcracker", "核桃夹"],
    "杯": ["tasse", "cup", "mug", "butelka", "kanne", "dzbanek", "kettle", "tumbler",
         "杯", "水壶", "保温"],
    "沙发套": ["sofabezug", "couch cover", "沙发套"],
    "门密封条": ["uszczelka", "door seal", "密封条"],
    "水龙头套": ["osłona na kran", "water faucet cover", "水龙头套", "wasserhahn"],
    "地漏": ["filtr odpływowy", "drain cover", "地漏"],
    # ---- 灯具电工 ----
    "小夜灯": ["nachtlicht", "lampka nocna", "night light", "夜灯"],
    "壁灯": ["wandleuchte", "kinkiet", "wall lamp", "壁灯"],
    "台灯": ["lampa stołowa", "desk lamp", "nachttischlampe", "台灯"],
    "吊灯": ["pendant", "żyrandol", "吊灯"],
    "头灯": ["czołowa", "headlamp", "head torch", "头灯"],
    "手电筒": ["latarka", "taschenlampe", "flashlight", "torch", "手电筒", "手电"],
    "工作灯": ["lampa robocza", "work light", "robocza", "工作灯"],
    "投光灯": ["außenstrahler", "reflektor", "flood light", "spot light", "投光", "射灯"],
    "氛围投影灯": ["projektor", "projector", "gwiazd", "northern lights", "starsky",
               "sternenlaser", "投影灯", "星空灯", "月球灯", "księżycowa"],
    "圣诞树灯": ["weihnachtsprojektor", "lampki", "choinkowe", "christmas projector", "lamki choinkowe", "fairy lights",
              "solar garden light", "christmas candle", "树顶", "tree topper", "urn filler",
              "zdoby choinkowe", "圣诞灯", "圣诞装饰", "santa hat"],
    "风扇": ["wentylator", "desk fan", "fan ", "电扇", "usb desk"],
    "插座电工": ["gniazdk", "outlet", "adapter", "插座"],
    # ---- 取暖制冷 ----
    "取暖器": ["heizgerät", "heizer", "grzejnik", "elektroheizer", "nagrzewnica", "nawiew",
            "heater", "取暖", "烤火炉", "molekularny"],
    "车载暖风机": ["auto heizung", "heiz- und kühl", "nagrzewnica samochodowa", "nagrzewnica",
             "ogrzewacz i chłodzenie", "chłodzenie samochodowe", "车暖", "车用暖风", "entfroster"],
    "车载除冰": ["eiskratzer", "ice scraper", "除冰"],
    "汽车遮雪挡": ["windschutzscheibenabdeckung", "windshield cover", "遮雪"],
    "暖手宝": ["handwärmer", "ogrzewacz do rąk", "hand warmer", "暖手"],
    "热水袋": ["hot water bottle", "wärmflasche", "热水袋"],
    # ---- 厨房 ----
    "切菜器": ["krajalnica", "gemuseschneider", "vegetable slicer", "mandolina", "切菜"],
    "煮蛋器": ["jajek na parze", "egg cooker", "煮蛋"],
    "净水器": ["filtr wody", "water filter", "净水"],
    "咖啡具": ["espresso", "coffee", "kawa", "咖啡"],
    "厨房工具": ["kitchen gadget", "przybory kuchenne", "厨具"],
    "餐具": ["sztućce", "cutlery", "plate", "talerz", "碗", "盘"],
    "保鲜盒": ["pojemnik", "food container", " Lunch box", "饭盒"],
    "围裙": ["schürze", "apron", "围裙"],
    "手套 oven": ["rękawice kuchenne"],
    # ---- 清洁 ----
    "吸尘器": ["staubsauger", "vacuum", "odkurzacz", "吸尘"],
    "拖把扫把": ["mop", "szczotka", "broom", "拖把", "扫把"],
    "洗衣": ["pralka", "laundry", "洗衣"],
    "垃圾袋": ["worek", "trash bag", "垃圾袋"],
    # ---- 工具五金 ----
    "电钻": ["bohrmaschine", "wiertarka", "drill", "电钻"],
    "电动螺丝刀": ["elektryczny srub", "screwdriver", "螺丝刀"],
    "电烙铁": ["lutownica", "soldering iron", "烙铁"],
    "点焊机": ["zgrzewarka", "spot welder", "点焊"],
    "万用表": ["multimetr", "multimeter", "万用"],
    "工具套装": ["werkzeug", "narzędzi", "tool kit", "tool set", "工具套装"],
    "刀具": ["nóż", "knife", "topór", "axe", "刀", "斧"],
    "焊锡钳子": ["szczypce", "pliers", "钳"],
    "梯凳": ["drabina", "ladder", "stool", "梯"],
    "胶水": ["klej", "glue", "adhesive", "胶"],
    "皮革修复": ["leather repair", "repair gel", "lederreparatur", "lederpflege",
             "żel naprawczy", "żel do skóry", "naprawczy do skóry", "naprawy skóry",
             "krem do naprawy", "皮革修复", "flicken", "latka"],
    "瓷砖修复": ["tile repair", "瓷砖修复"],
    "家具滑轮": ["möbelrollen", "furniture glider", "滑轮"],
    "家具搬运带": ["möbeltragegurt", "moving straps", "搬运带"],
    "家具护理": ["möbelpflege", "pielęgnacji mebli", "皮革保养", "家具护理"],
    "门锁": ["türschloss", "zamek", "door lock", "schloss", "门锁", "密码锁"],
    # ---- 车载 ----
    "行车记录仪": ["dash cam", "autokamera", "行车记录"],
    "车载HUD": ["autotacho", "hud", "抬头显示"],
    "车载支架": ["uchwyt samochodowy", "phone mount", "car holder", "手机支架"],
    "车载充电": ["ładowarka samochodowa", "car charger", "车充"],
    "自行车配件": ["fahrrad", "rower", "bicycle", "bike", "自行车"],
    "安全反光衣": ["sicherheitsweste", "hi-vis", "reflective vest", "反光衣"],
    # ---- 3C电子 ----
    "蓝牙耳机": ["kopfhörer", "słuchawki", "earbuds", "headphones", "headset", "耳机"],
    "睡眠耳机": ["sleep headphones", "opaska na głowę ze słuchawk", "睡眠耳"],
    "音箱": ["głośnik", "speaker", "bluetooth box", "音箱"],
    "麦克风": ["mikrofon", "microphone", "话筒"],
    "K歌机": ["karaoke", "karaoke maschine", "karaoke machine", "mikrofon do karaoke", "k歌"],
    "收音机": ["hand crank radio", "radio", "kurbel", "notradio", "收音机"],
    "对讲机": ["walkie", "interkom", "对讲"],
    "数码相机": ["aparat", "digital camera", "相机"],
    "摄像头": ["überwachungskamera", "surveillance", "security camera", "monitoring",
            "mini kamera", "kamera", "camera", "摄像头"],
    "数码相机": ["aparat", "digital camera", "相机"],
    "WiFi放大器": ["wzmacniacz wi-fi", "wifi extender", "wifi repeater", "信号放大"],
    "充电宝": ["solar power bank", "power bank", "powerbank", "充电宝"],
    "充电器": ["ładowarka", "charger", "usb ladegerät", "充电"],
    "智能手表": ["smartwatch", "zegrek", "手表"],
    "电子秤": ["waga", "scale", "skala", "秤"],
    "计算器": ["kalkulator", "calculator", "计算"],
    "激光笔": ["wskaznik laserowy", "laser pointer", "激光笔"],
    "放大镜": ["lupa", "lupę", "szkło powiększające", "magnifying", "vergrößer", "放大镜"],
    "望远镜": ["lornetka", "binoculars", "monokular", "monocular", "fernglas", "teleskop", "望远镜"],
    "指南针": ["kompass", "compass", "指北", "指南针"],
    "闹钟": ["budzik", "alarm clock", "闹钟"],
    "验钞": ["banknot", "money detector", "验钞"],
    # ---- 健康护理 ----
    "矫姿带": ["korektor postawy", "haltungskorrektor", "posture corrector", "矫姿"],
    "护膝": ["kolano", "knie", "knee", "kolanowego", "kolana", "护膝", "膝盖"],
    "护肩": ["barku", "shoulder brace", "肩", "护肩"],
    "护肘": ["okieć", "elbow", "护肘"],
    "护腰带": ["pas lędźwiowy", "lumbar", "nieren", "护腰", "托腹带", "sciatic", "坐骨", "schwangerschaftsgürtel"],
    "压力袜": ["kompressionsstrümpfe", "compression socks", "uciskowe", "skarpety uciskowe", "压力袜"],
    "拇外翻矫正": ["bunion", "halux", "hallux", "拇外翻"],
    "扁平足鞋垫": ["wkładki ortopedyczne", "płaskostopie", "flat feet", "足弓"],
    "止鼾器": ["chrpani", "chrapani", "anti-snoring", "止鼾"],
    "助行器": ["standing aid", "walker", "助行"],
    "按摩器": ["massager", "masażer", "massage", "masujący", "按摩"],
    "筋膜枪": ["massage gun", "pistola do masażu", "筋膜"],
    "热敷眼罩": ["maska grzewcza na oczy", "eye mask", "heat eye", "眼罩"],
    "睡眠眼罩": ["maska do spania", "masce do spania", "sleep mask", "blaspheme", "遮光眼罩", "眼罩睡眠"],
    "助眠枕": ["reflux pillow", "反流枕"],
    "加热护具": ["grzewcze", "ogrzewane", "beheizbar", "unterwaerm", "heated", "加热"],
    "暖宫腰带": ["menstru", "do bólu menstruacyjnego", "月经", "暖宫"],
    "体温计": ["thermometer", "termometr", "体温"],
    "指甲刀": ["nail clipper", "do paznokci", "nagel", "指甲"],
    "理发器": ["trymer", "trimmer", "klipper", "haarschneide", "理发"],
    "剃须刀": ["rasierer", "razor", "gilotyna", "剃须"],
    "牙刷架": ["toothbrush holder", "zahnborstenhalter", "uchwyt na szczoteczki", "szczoteczki do zębów", "牙刷卡座"],
    "电动牙刷": ["sonic toothbrush", "szczoteczka", "电动牙刷"],
    "金属探测器": ["metal detector", "detektor metali", "gold and treasure", "金属探测", "探金"],
    "猫抓板": ["scratching", "cat scratcher", "drapak", "猫抓"],
    "磨粉机": ["grain mill", "flour grinder", "młynek", "磨粉机", "谷物研磨"],
    "冲牙器": ["irrigator", "water flosser", "冲牙", "płukania jamy"],
    "头皮按摩器": ["scalp massager", "kopfs", "头皮"],
    "卷发工具": ["lockenwickler", "curler", "hair curl", "卷发"],
    "发膜": ["haarmaske", "hair mask", "maska do włosów", "kur do włosów", "发膜", "焗油"],
    "头发护理": ["hair spray", "spray do włos", "serum", "shampoo", "护发", "洗发"],
    "干发帽": ["czepek na włosy", "dry hair cap", "干发"],
    "面霜": ["krem do twarzy", "face cream", "cf", "面霜", "nickotynamid"],
    "驱蚊手环": ["opaska przeciw owadom", "mosquito band", "驱蚊手环"],
    "灭蚊灯": ["fly trap", "mosquito zapper", "insect killer", "mücken", "灭蚊", "捕蚊"],
    "灭鼠器": ["myszy", "ratten", "mouse trap", "rat trap", "捕鼠", "灭鼠"],
    "射击耳罩": ["gehörschutz", "schießsport", "耳罩降噪"],
    "激光笔": ["wskaźnik", "wskaznik"],
    "感应灯": ["czujnik ruchu", "czujnikiem ruchu", "感应灯"],
    "防噎急救": ["zadławieniu", "heimlich", "防噎"],
    "急救包": ["first aid", "apteczka", "急救"],
    "药盒": ["pillenorganizer", "pill box", "tygodniowa", "药盒"],
    # ---- 母婴儿童 ----

    "婴儿背带": ["nosidełko", "baby carrier", "背带"],
    "儿童坐便器": ["potty", "toilet seat", "training", "坐便"],
    "安全门栏": ["travel gate", "barierka", "门栏"],
    "蒙氏教具": ["tablica sensoryczna", "sensory board", "sensoryczna", "早教板"],
    "积木": ["klocki", "building blocks", "lego", "积木"],
    "打击乐": ["schlagzeug", "perkusyjny", "drum", "ksylofon", "打击乐", "打击乐器"],
    "益智玩具": ["fidget", "puzzle", "strategy board", "dice game", "spinner", "学习玩具",
             "math游戏", "mathe", "益智"],
    "玩具车": ["rc car", "rc auto", "rc-auto", "remote controlled car", "remote control car",
           "ferngesteuert", "ferngesteuertes auto", "zdalnie sterowany", "spielzeugauto",
           "samochód", "excavator", "koparka", "kletterauto", "transformer", "robot",
           "遥控车", "遥控汽车", "toyota"],
    "遥控船": ["rc boat", "remote control boat", "łódź", "boat", "遥控船"],
    "遥控直升机": ["rc helicopter", "hubschrauber", "helicopter", "gyroskop", "遥控直升机"],
    "玩具飞机": ["spielzeugflugzeug", "flying aeroplane", "aeroplane", "flugzeug",
           "plane toy", "airplane toy", "玩具飞机"],
    "方向盘玩具": ["lenkrad", " steering wheel toy", "方向盘"],
    "逗猫玩具": ["cat toy", "flapping feather", "interactive cat", "逗猫"],
    "玩具火车": ["eisenbahn", "train toy", "kolejka", "火车"],
    "秋千": ["schaukel", "cocoon swing", "sensory swing", "吊椅", "秋千"],
    "戏水玩具": ["splash mat", "water slide", "piaskownica", "water and sand", "play table",
            "戏水", "滑梯"],
    "睡衣": ["piżama", "pizama", "schlafanzug", "pyjama", "睡衣"],
    "演出服": ["kostüm", "cosplay", "halloween", "santa claus", "vampir", "服装演出", "cos"],
    # ---- 宠物 ----
    "狗胸背带": ["geschirr", "harness", "szelki dla psa", "szelki", "uprząż", "胸背"],
    "宠物窝": ["legowisko", "pet bed", "kennel", "猫窝", "狗窝"],
    "喂食器": ["feeder", "karmnik", "fountain", "trink", "自动喂", "喂食"],
    "喂鸟器": ["bird feeder", "domek dla ptaków", "karmnik dla ptaków", "karmnik", "Vogelhaus", "鸟屋"],
    "止吠器": ["bark stopper", "antyszczek", "止吠"],
    "训狗项圈": ["halsband", "dog collar", "elektrisches hundehalsband", "项圈"],
    "宠物耳罩": ["hearing protectors for", "dog ear", "宠物耳"],
    "宠物剃毛器": ["trymer do sierści", "maszynka do strzyżenia", "pet clipper", "pet trimmer", "宠物剃"],
    "宠物洁牙": ["dental", " Teeth Cleaning Device for Canine", "dog tooth", "宠物牙"],
    "驱虫驱鸟": ["bird repeller", "bird deterrent", "ultraschall", "ultrasonic", "zwierzęta odstrasz",
             "repeller", "deterrent", "驱鸟", "驱虫"],
    "拾便器": ["zbieracz", "poop", "拾便"],
    # ---- 户外园艺 ----
    "吊床": ["hammock", "吊床"],
    "帐篷": ["zelt", "namiot", "tent", "帐篷"],
    "睡袋": ["sleeping bag", "śpiwór", "睡袋"],
    "登山杖": ["hiking poles", "trekking", "krkije", "登山杖"],
    "野餐": ["picnic", "plażowa", "matę", "沙滩垫"],
    "水管": ["garden hose", "wąż", "水管", " ogrod"],
    "洒水器": ["sprinkler", "raszcz", "洒水"],
    "园艺工具": ["gartengerät", "garden tool", "scraper for weeds", "grubka", "除草", "园艺"],
    "花盆": ["doniczka", "planter", "vase", "花盆"],
    "泡泡机": ["bubble machine", "bańki", "泡泡"],
    "风筝": ["drachen", "latawiec", "kite", "风筝"],
    "戏水喷泉": ["springwater", "water fountain", "戏水"],
    # ---- 洗漱浴室 ----
    "花洒": ["duschkopf", "prysznic", "shower head", "花洒"],
    "浴巾架": ["handtuchhalter", "ręcznik holder", "毛巾架"],
    "牙刷架": ["toothbrush holder", "zahnborstenhalter", "uchwyt na szczoteczki", "szczoteczki do zębów", "牙刷卡座"],
    "肥皂盒": ["seifendose", "soap dish", "肥皂"],
    # ---- 其他 ----
    "雨伞": ["regenschirm", "parasol", "umbrella", "雨伞"],
    "雨靴": ["stiefel wasserdicht", "rain boots", "gummi", "雨靴"],
    "存钱罐": ["skarbonka", "money box", "piggy", "存钱罐"],
    "香薰机": ["diffusor", "aroma", "humidifier", "nawilżacz", "香薰", "加湿"],
    "蜡烛": ["kerze", "swieca", "candle", "蜡烛"],
    "相框": ["foto", "ramka", "obraz", "相框", "画框"],
    "墙贴": ["aufkleber", "wall sticker", "naklejka", "墙贴"],
    "门帘": ["door curtain"],
    "钥匙扣": ["keychain", "brelok", "karabielski", "钥匙"],
    "笔": ["pen", "długopis", "stift", "pióro", "笔"],
    "笔记本": ["notes", "zeszyt", "notbook", "笔记本"],
    "背包挂饰": ["charm", "breloczek", "挂饰"],
    "3D打印笔": ["3d pen", "3d-stift", "dlugopis 3d", "długopis 3d", "gravierstift", "3d打印", "雕刻笔", "do grawerowania"],
    "熨烫": ["bryczes", " żelazko", "iron", "熨斗"],
    "缝纫": ["nähmaschine", "maszyna do szycia", "sewing", "缝纫"],
    "毛球修剪器": ["golarka do kłaczków", "lint remover", "fabri", "毛球"],
    "加湿器": ["luftbefeuchter", "humidifier", "加湿器"],
    "除湿": ["dehumidifier", "osuszacz", "除湿"],
    "净化器": ["luftreiniger", "air purifier", "oczyszczacz", "净化"],
    "电风扇": ["klimatyzator", "air conditioner", "空调扇"],
    "洗衣凝珠": ["kapsułki do prania"],
    "晾衣架": ["stang", "wieszak na ubrania", "晾衣"],
    "熨衣板": ["bügelbrett", "ironing board"],
    "烫发器": ["prostownica", "flat iron", "styl", "夹板"],
    "化妆镜": ["lusterko", "makeup mirror", "化妆镜"],
}


def cn_of(text):
    """标题 → 中文品名。取「最长匹配词」所在品类（具体词天然赢过宽泛词），
    长度相同则按词表顺序优先。未识别返回 ''。"""
    t = re.sub(r"[\s\u00a0\u202f]+", " ", str(text or "")).lower().replace("–", "-")
    if not t:
        return ""
    best, best_len = "", 0
    for cat, words in CATEGORY_WORDS.items():
        for w in words:
            if len(w) > best_len and w in t:
                best, best_len = cat, len(w)
    return best


_CN_MEMO = {}


def cn_of_names(names):
    """一组标题里逐个试，返回第一个识别出的中文品名（memo 缓存加速）。"""
    key = "\x01".join(str(n or "")[:120] for n in (names or [])[:4])
    if not key:
        return ""
    hit = _CN_MEMO.get(key)
    if hit is not None:
        return hit
    c = ""
    for n in names or []:
        c = cn_of(n)
        if c:
            break
    _CN_MEMO[key] = c
    return c


CN_FILE = os.path.join(BASE, "cn_names.json")
_CN_OVERRIDE = None


def _load_overrides():
    global _CN_OVERRIDE
    if _CN_OVERRIDE is None:
        try:
            if os.path.exists(CN_FILE):
                _CN_OVERRIDE = json.load(open(CN_FILE, encoding="utf-8"))
            else:
                _CN_OVERRIDE = {}
        except Exception:
            _CN_OVERRIDE = {}
    return _CN_OVERRIDE


def cn_for(p, url=None):
    """产品/配对/新品 dict → 中文品名。
    优先「看图核对表」cn_names.json（人工+图片校准），其次标题词表推导。"""
    ov = _load_overrides()
    for u in (norm(url or ""), norm(p.get("url") or ""), p.get("inq") or ""):
        if u and u in ov:
            return ov[u]
    return cn_of_names(p.get("names") or [p.get("name", "")])


def _cats(text):
    """返回文本命中的品类集合。"""
    t = str(text).lower()
    out = set()
    for cat, words in CATEGORY_WORDS.items():
        for w in words:
            if w in t:
                out.add(cat)
                break
    return out


def cat_overlap(a_names, b_names):
    """两组标题的品类交集是否有重合。"""
    ca = set()
    for n in a_names:
        ca |= _cats(n)
    cb = set()
    for n in b_names:
        cb |= _cats(n)
    return ca & cb, ca, cb


def _spec_tokens(text):
    """抽取标题里的规格参数：尺寸(50x138)、容量(mAh/L/ml)、功率(W)、数量等。"""
    t = str(text or "").lower().replace("×", "x").replace("*", "x")
    out = set()
    # 尺寸 50x138 / 90x45
    for a, b in re.findall(r"(\d{2,3})\s*x\s*(\d{2,3})", t):
        out.add("%sx%s" % (a, b))
    # 带单位的数值
    for v, u in re.findall(r"(\d+(?:[.,]\d+)?)\s*(mm|cm|ml|l|w|mah|kg|g|v|inch|in|zoll)\b", t):
        v = v.replace(",", ".")
        out.add(v + u)
    return out


def find_same_by_title_spec(target_names, pool, exclude_url=None):
    """标题规格参数兜底：像「皮革修补贴 50x138」这种，同款不同配色时图片差异大，
    但标题里的规格数字一致 + 品名词一致，就基本可确定是同一款。"""
    tname = (target_names or [""])[0]
    tspec = _spec_tokens(tname)
    if not tspec:
        return None, None
    tcn = cn_of_names(target_names)
    tu = norm(exclude_url or "")
    best, best_info = None, None
    for x in pool:
        if tu and norm(x.get("url")) == tu:
            continue
        xs = set()
        for n in (x.get("names") or [""])[:3]:
            xs |= _spec_tokens(n)
        common = tspec & xs
        if not common:
            continue
        # 规格必须覆盖主要参数（至少一个，且不能只是很小的巧合）
        if len(common) < 1:
            continue
        xcn = cn_of_names(x.get("names") or [""])
        # 品名相同 → 强证据；品名不同则要求 2 个以上规格重合
        if tcn and xcn and tcn == xcn:
            score = 2
        elif len(common) >= 2:
            score = 1
        else:
            continue
        if best is None or score > best_info["score"]:
            best, best_info = x, {"spec": sorted(common), "cn": xcn, "score": score}
    return best, best_info


def find_same_by_image(target_img_url, pool, exclude_url=None):
    """按图片判定库内同款（跨站点、不同语言标题的场景）。
    严格三重条件——宁可漏判，也绝不误带 SPU：
      1) 精细图距 fine <= 320
      2) 色相相似 hue >= 0.80（真同款基本 >0.95，纯白底图乱配通常 <0.6）
      3) 双方都是纯白底图时，还要求亮度/结构也足够接近
    返回 (候选dict, 距离信息) 或 (None, None)。"""
    tsig, _ = sig_for_url(target_img_url)
    if not tsig:
        return None, None
    thue = _hue_of(tsig)
    twhite = _white_of(tsig)
    tu = norm(exclude_url or "")
    best_x, best_df, best_info = None, 99999, None
    for x in pool:
        if tu and norm(x.get("url")) == tu:
            continue
        xsig, _ = sig_for_url(x.get("img"))
        if not xsig:
            continue
        df = fine_dist(tsig, xsig)
        if df >= best_df or df > 320:
            continue
        hc = cosv(thue, _hue_of(xsig))
        if hc < 0.80:
            continue
        d, ds = hdist(tsig, xsig), sdist(tsig, xsig)
        xwhite = _white_of(xsig)
        if twhite > 0.60 and xwhite > 0.60 and d > 90 and ds > 90:
            continue
        best_x, best_df = x, df
        best_info = {"fine": df, "hue": round(hc, 3), "h": d, "s": ds}
    return best_x, best_info


def find_similar(target_names, target_img_url, pool, topn=3, min_score=0.50,
                 exclude_inq=None, exclude_rows=None, exclude_spu=None, t_url=None):
    """
    在 pool 中找相似款。
    pool: [{"row","names","url","inq","spu","spus","img","u1688"}]
    exclude_inq: 要排除的询盘号（即目标自身）
    返回按匹配度排序的候选列表。
    """
    tname = target_names[0] if target_names else ""
    tsig, tpath = sig_for_url(target_img_url)
    thue = _hue_of(tsig)
    twhite = _white_of(tsig) if tsig else 0.0
    turl = norm(target_img_url)
    tnames_set = set(target_names)
    tcn = cn_for({"names": target_names}, url=t_url)   # 目标中文品名（审核表优先）

    scored = []
    for x in pool:
        # 排除自身：同链接 / 同询盘号 / 标题完全相同 / 指定行号
        if turl and norm(x.get("url")) == turl:
            continue
        if exclude_inq and x.get("inq") and x.get("inq") == exclude_inq:
            continue
        if exclude_rows and x.get("row") in exclude_rows:
            continue
        # 同 SPU = 同一产品，不再作为候选
        if exclude_spu:
            xs = set(x.get("spus") or ([x["spu"]] if x.get("spu") else []))
            if exclude_spu & xs:
                continue
        # 同一条记录（主标题相同）才算自己；多语言 names 数组有交集不代表是同一行
        if x.get("names") and tname and x["names"][0] == tname:
            continue
        ns = max((nsim(tname, n) for n in x.get("names", [])), default=0.0)
        xsig, xpath = sig_for_url(x.get("img"))
        if xsig is None and xpath is not None:
            xsig = img_sig(xpath)
        if tsig is None and tpath is not None:
            tsig = img_sig(tpath)
        d = hdist(tsig, xsig)
        ds = sdist(tsig, xsig)
        dp = shapdist(tsig, xsig)
        df = fine_dist(tsig, xsig)
        # 图片几乎相同 = 同款，不受标题语言/中文品名归类影响（跨语言同款的关键兜底）
        img_tiny = (tsig is not None and xsig is not None
                    and (d <= 60 or ds <= 70 or dp <= 60))
        # ---- 中文品名粗筛：同名才精比；跨品类/未识别类目必须标题够像或图片极近 ----
        xcn = x.get("cn") or cn_for(x)
        if tcn and not img_tiny:
            if xcn and xcn != tcn and ns < 0.62:
                continue
            # 认不出品类（多为其它语种的同款变体）时不硬过滤，交给图片与分数判定
        hc = cosv(thue, _hue_of(xsig)) if xsig else 0.0
        xwhite = _white_of(xsig) if xsig else 0.0
        # 双方都是白底图时，色相不可靠 → 置中性值
        if twhite > 0.45 and xwhite > 0.45:
            hc = 0.0
        has_img = xsig is not None and tsig is not None
        # 有效距离：16x16 与 32x32 取优
        d_eff = min(d, ds, dp, df // 6) if has_img else 9999

        # 跨语言品类识别
        ov, ca, cb = cat_overlap(target_names, x.get("names", []))

        # 实词重合计分：多语言标题里能共现的实词越多，越可能是同一款（跳过 stop 词）
        tk, xk = toks(tname), toks(x.get("names", [""])[0])
        shared_kw = len(tk & xk)

        # ===== 多证据融合判定（经 8组正样本 + 200组负样本标定）=====
        #   same     : 16x16极小 或 (精细距<=600 且 色相>=0.75)
        #   likely   : (精细<=750 且 色相>=0.55)
        #   possible : (精细<=850 且 色相>=0.50)
        tiny = has_img and (d <= 40 or ds <= 55 or dp <= 45)
        # 白底图上色相不可靠 → 必须靠强图距或品类
        white_pair = twhite > 0.45 and xwhite > 0.45

        if tiny:
            level = "same"
        elif has_img and df <= 600 and hc >= 0.75 and not white_pair:
            level = "same"
        elif has_img and df <= 600 and ov and ns >= 0.30:
            level = "same"                          # 图近 + 同品类 + 标题沾边
        elif has_img and xcn and tcn and xcn == tcn and df <= 600:
            level = "same"                          # 同一中文品名 + 图片高度相近（跨语种同款）
        elif has_img and df <= 700 and hc >= 0.55 and not white_pair:
            level = "likely"
        elif has_img and df <= 750 and ov and ns >= 0.25:
            level = "likely"
        elif has_img and xcn and tcn and xcn == tcn and df <= 750:
            level = "likely"                        # 同一中文品名 + 图片相近
        elif has_img and df <= 800 and hc >= 0.50 and not white_pair:
            level = "possible"
        elif has_img and df <= 850 and ov and ns >= 0.20:
            level = "possible"
        elif has_img and xcn and tcn and xcn == tcn and df <= 900:
            level = "possible"                      # 同一中文品名兜底：交给人工确认
        elif has_img and xcn and tcn and xcn == tcn and shared_kw >= 2 and df <= 1050:
            level = "possible"                      # 同品名 + 多个实词重合（不同语言/摆拍图也能捞回）
        elif has_img and xcn and tcn and xcn == tcn and hc >= 0.78 and df <= 1050:
            level = "possible"                      # 同品名 + 主色一致（摆拍 vs 实拍图的兜底）
        elif ns >= 0.78:                            # 标题几乎相同
            level = "likely"
        elif ns >= 0.65 and ov:
            level = "possible"
        else:
            continue

        # 综合分（含品类加分）
        if has_img:
            fine_score = max(0.0, (1200 - min(df, 1200)) / 1200.0)
            small_score = max(0.0, (300 - min(d_eff, 300)) / 300.0)
            score = (0.40 * fine_score + 0.25 * hc + 0.20 * small_score
                     + 0.10 * ns + (0.05 if ov else 0.0))
        else:
            score = ns * 0.75 + (0.25 if ov else 0.0)

        # 同中文品名的候选放宽门槛：属于同一类目、图和标题都还沾边，值得人工看一眼
        same_cn = bool(tcn and xcn and tcn == xcn)
        eff_min = min_score if not same_cn else min(min_score, 0.42)
        if score < eff_min and level not in ("same", "likely"):
            continue
        if level == "possible" and score < eff_min:
            continue
        scored.append({
            "row": x.get("row"),
            "name": x.get("names", [""])[0],
            "names": x.get("names", []),
            "cn": xcn,
            "url": x.get("url", ""),
            "inq": x.get("inq", ""),
            "spu": x.get("spu", ""),
            "spus": x.get("spus", []),
            "img": x.get("img", ""),
            "dist": d_eff if has_img else None,
            "dist_bright": d if has_img else None,
            "dist_struct": ds if has_img else None,
            "dist_shape": dp if has_img else None,
            "dist_fine": df if has_img else None,
            "hue": round(hc, 3),
            "nsim": round(ns, 2),
            "score": round(score, 3),
            "level": level,
        })
    scored.sort(key=lambda z: -z["score"])
    _save_sigs()
    return scored[:topn]


LEVEL_TXT = {"same": "高度相似", "likely": "可能相似", "possible": "略有相似"}
