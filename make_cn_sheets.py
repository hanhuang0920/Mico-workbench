# -*- coding: utf-8 -*-
"""生成中文品名看图核对表：每页 20 个产品图 + 编号 + 当前名。
用法: python make_cn_sheets.py   → 输出到 cn_sheets/sheet_NN.png + cn_sheets/index.json
"""
import os
import sys
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from PIL import Image, ImageDraw, ImageFont
import supabase_store as sb
import similar

OUT = os.path.join(similar.BASE, "cn_sheets")
CELL_W, IMG_H, LABEL_H = 260, 200, 88
COLS, ROWS = 5, 4
PAD = 10


def font(size):
    for cand in ("msyh.ttc", "msyhbd.ttc", "segoeui.ttf", "arial.ttf"):
        try:
            return ImageFont.truetype(cand, size)
        except Exception:
            continue
    return ImageFont.load_default()


def main():
    os.makedirs(OUT, exist_ok=True)
    prods = sb.get_products()
    index = []
    tiles = []
    f14 = font(14)
    f12 = font(12)
    for i, p in enumerate(prods):
        url = p.get("url") or ""
        img_url = p.get("img") or ""
        name = (p.get("names") or [""])[0]
        cn = similar.cn_for(p)
        path = similar.ensure_img(img_url)
        tile = Image.new("RGB", (CELL_W, IMG_H + LABEL_H), "white")
        dr = ImageDraw.Draw(tile)
        if path:
            try:
                im = Image.open(path).convert("RGB")
                im.thumbnail((CELL_W - 8, IMG_H - 8))
                tile.paste(im, (4, 2))
            except Exception:
                dr.text((10, 80), "IMAGE ERROR", fill="red", font=f14)
        else:
            dr.text((10, 80), "NO IMAGE", fill="gray", font=f14)
        # 标签: 编号 | 当前名 | url短 | inq
        short_url = url.replace("https://", "").replace("http://", "")
        if len(short_url) > 34:
            short_url = short_url[:34] + "…"
        dr.rectangle([0, IMG_H, CELL_W, IMG_H + LABEL_H], fill="#f1f5f9")
        dr.text((6, IMG_H + 2), "#%d  %s" % (i, cn or "未识别"), fill="#b91c1c" if not cn else "#0f172a", font=f14)
        dr.text((6, IMG_H + 24), name[:40], fill="#334155", font=f12)
        dr.text((6, IMG_H + 42), name[40:80], fill="#334155", font=f12)
        dr.text((6, IMG_H + 60), short_url, fill="#64748b", font=f12)
        dr.text((6, IMG_H + 75), (p.get("inq") or "")[-14:], fill="#94a3b8", font=f12)
        tiles.append(tile)
        index.append({"i": i, "url": url, "inq": p.get("inq") or "",
                      "cur_cn": cn, "name": name})

    per = COLS * ROWS
    sheets = (len(tiles) + per - 1) // per
    W = COLS * (CELL_W + PAD) + PAD
    H = ROWS * (IMG_H + LABEL_H + PAD) + PAD
    for s in range(sheets):
        sheet = Image.new("RGB", (W, H), "#e2e8f0")
        for j in range(per):
            k = s * per + j
            if k >= len(tiles):
                continue
            r, c = divmod(j, COLS)
            sheet.paste(tiles[k], (PAD + c * (CELL_W + PAD), PAD + r * (IMG_H + LABEL_H + PAD)))
        fp = os.path.join(OUT, "sheet_%02d.png" % (s + 1))
        sheet.save(fp)
    json.dump(index, open(os.path.join(OUT, "index.json"), "w", encoding="utf-8"), ensure_ascii=False)
    print("sheets:", sheets, "tiles:", len(tiles), "→", OUT)


if __name__ == "__main__":
    main()
