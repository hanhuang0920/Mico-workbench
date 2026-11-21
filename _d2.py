# -*- coding: utf-8 -*-
"""并排导出两条产品的图片，人工核对是否同款"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import supabase_store as sb
import similar
from PIL import Image

OUT = os.path.join(similar.BASE, "_cmp")
os.makedirs(OUT, exist_ok=True)

prods = sb.get_products()
news = sb.get_news()

targets = []
for p in prods:
    nm = (p.get("names") or [""])[0]
    if "Baby Pillow for Reflux" in nm:
        targets.append(("A_产品库_BabyPillow", p))
for p in news:
    nm = (p.get("names") or [p.get("name", "")])[0]
    if "Miękka Poduszka" in nm:
        targets.append(("B_新品_MiekkaPoduszka", p))

imgs = []
for tag, p in targets:
    path = similar.ensure_img(p.get("img"))
    if not path:
        print("下载失败:", tag)
        continue
    im = Image.open(path).convert("RGB")
    im.thumbnail((420, 420))
    imgs.append((tag, im))
    print("%s  %s  size=%s" % (tag, (p.get("names") or [p.get("name", "")])[0][:50], im.size))

if len(imgs) >= 2:
    W = sum(i[1].width for i in imgs) + 20 * (len(imgs) + 1)
    H = max(i[1].height for i in imgs) + 20
    canvas = Image.new("RGB", (W, H), "white")
    x = 20
    for tag, im in imgs:
        canvas.paste(im, (x, 10))
        x += im.width + 20
    fp = os.path.join(OUT, "baby_pillow.png")
    canvas.save(fp)
    print("拼图:", fp)
