# -*- coding: utf-8 -*-
"""相似款匹配引擎：图片哈希 + 标题相似度，返回候选及 SPU。"""
import sys, io, os, re, json, urllib.request
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from difflib import SequenceMatcher

BASE = r"D:\DSH\Mico"
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
_URL_SIGS = None


def _load_url_sigs():
    global _URL_SIGS
    if _URL_SIGS is None:
        _URL_SIGS = {}
        try:
            import supabase_store as _sb
            if _sb.enabled():
                _URL_SIGS = _sb.get_img_sigs()
        except Exception as e:
            print("[img_sigs 读取失败，退回在线下载] %s" % e)
    return _URL_SIGS


def sig_for_url(url):
    """优先用 Supabase 签名缓存（不下载图片）；缺失才下载计算并写回。
    返回 (sig, path_or_None)：path 为 None 表示无需本地文件。"""
    u = norm(url)
    if not u or not str(url).startswith("http"):
        return None, None
    sigs = _load_url_sigs()
    if u in sigs and len(sigs[u]) >= 9:
        return tuple(sigs[u]), None
    p = ensure_img(url)
    sig = img_sig(p) if p else None
    if sig is not None:
        sigs[u] = sig
        try:
            import supabase_store as _sb
            if _sb.enabled():
                _sb.put_img_sig(u, [x if not isinstance(x, tuple) else list(x) for x in sig])
        except Exception:
            pass
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
    """形状距离：仅比对裁剪归一化后的轮廓，对尺寸/裁剪/颜色都鲁棒。"""
    if not a or not b or len(a) < 5 or len(b) < 5:
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


# ===== 跨语言产品类型词表 =====
CATEGORY_WORDS = {
    "羊绒衫": ["kaschmir pullover", "kaszmir", "kaschmirpullover", "cashmere sweater",
              "sweter meski", "sweter", "pullover", "strickpullover", "羊绒衫"],
    "打底裤": ["legginsy", "leggings", "thermo leggings", "winterleggings", "打底裤"],
    "派克大衣": ["parka", "winterparka", "wiatr", "winddicht", "派克"],
    "雨衣": ["regenjacke", "raincoat", "kurtka przeciwdeszczowa"],
    "围巾脖套": ["schlauchschal", "szalik", "komin", "scarf", "haarband", "围巾"],
    "秋千": ["kokonschaukel", "cocoon swing", "schaukel", "swing", "秋千"],
    "K歌机": ["karaoke", "k歌"],
    "睡衣": ["piżama", "pizama", "schlafanzug", "pyjama", "睡衣"],
    "收音机": ["hand crank radio", "radio awaryjne", "kurbel", "notradio", "收音机"],
    "投影灯": ["projektor", "projekt", "projector", "lichteffekt", "投影"],
    "望远镜": ["lornetka", "binoculars", "monokular", "fernglas", "望远镜"],
    "收音": ["radio"],
    "手套": ["rekawiczki", "handschuhe", "gloves", "rękawice", "手套"],
    "皮革贴": ["latka", "flicken", "repair patch", "皮革贴"],
    "切菜器": ["krajalnica", "gemuseschneider", "vegetable slicer", "slicer"],
    "3D笔": ["3d pen", "3d-stift", "dlugopis 3d", "gravierstift", "3d打印笔"],
    "遥控车": ["zdalnie sterowany", "remote control car", "ferngesteuert", "rc car"],
    "打击乐": ["schlagzeug", "perkusyjny", "drum"],
    "羊绒围巾": ["kaszmir", "kaschmir", "cashmere"],
    "门锁": ["turschloss", "zamek", "door lock", "密码锁"],
    "摄像头": ["kamera", "camera", "摄像头"],
    "加热袜": ["beheizbare socken", "skarpetki", "heated socks", "加热袜"],
    "弹簧摆件": ["kinetischer", "kinetic", "molekularheizer", "分子加热",
              "spring toy", "sprężyn", "feder", "弹簧", "摆件",
              "car perfume", "auto duft", "香薰摆件", "车载摆件"],
    "唇膏口红": ["lipstick", "lippenstift", "szminka"],
    "拖鞋": ["hausschuhe", "kapcie", "slippers", "拖鞋", "schlappen"],
    "保温杯": ["thermos", "tumbler", "butelka termiczna", "保温"],
    "书包背包": ["rucksack", "backpack", "plecak", "背包", "书包"],
    "蓝牙耳机": ["kopfhörer", "sluchawki", "earbuds", "headphones", "耳机"],
    "充电宝": ["power bank", "powerbank", "充电宝"],
    "电钻": ["bohrmaschine", "wiertarka", "drill", "电钻"],
    "工具套装": ["werkzeug", "narzędzi", "tool kit", "工具"],
    "地毯地垫": ["teppich", "dywan", "地毯", "地垫", "mat"],
    "窗帘": ["vorhang", "zasłona", "curtain", "窗帘"],
    "抱枕": ["kissen", "poduszka", "pillow", "抱枕", "靠垫"],
    "毛巾": ["handtuch", "ręcznik", "towel", "毛巾"],
    "雨伞": ["regenschirm", "parasol", "umbrella", "雨伞"],
    "保温壶": ["kanne", "dzbanek", "kettle", "水壶"],
    "收纳盒": ["aufbewahrung", "pudełko", "storage box", "收纳"],
    "书架": ["regal", "półka", "shelf", "书架"],
    "挂钩": ["haken", "haczyk", "hook", "挂钩"],
    "花洒": ["duschkopf", "prysznic", "shower head", "花洒"],
}


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


def find_similar(target_names, target_img_url, pool, topn=3, min_score=0.50,
                 exclude_inq=None, exclude_rows=None, exclude_spu=None):
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
        if x.get("names") and (set(x["names"]) & tnames_set):
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
        elif has_img and df <= 700 and hc >= 0.55 and not white_pair:
            level = "likely"
        elif has_img and df <= 750 and ov and ns >= 0.25:
            level = "likely"
        elif has_img and df <= 800 and hc >= 0.50 and not white_pair:
            level = "possible"
        elif has_img and df <= 850 and ov and ns >= 0.20:
            level = "possible"
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

        if score < min_score and level not in ("same", "likely"):
            continue
        if level == "possible" and score < min_score:
            continue
        scored.append({
            "row": x.get("row"),
            "name": x.get("names", [""])[0],
            "names": x.get("names", []),
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
