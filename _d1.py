# -*- coding: utf-8 -*-
"""诊断：这两条婴儿枕是否为同款却没比出来"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import supabase_store as sb
import similar

TAGS = ["Miękka Poduszka Dziecięca", "Baby Pillow for Reflux",
        "Poduszka dla Noworodka", "Poduszka do Leżenia", "Poduszka Pozycjonująca"]

def locate(pool, tag, label):
    for i, p in enumerate(pool):
        nm = (p.get("names") or [p.get("name", "")])[0]
        if tag.lower() in nm.lower():
            return i, p, label
    return None, None, label

prods = sb.get_products()
pairs = sb.get_pairs()
news = sb.get_news()
print("产品库 %d / 配对 %d / 新品 %d" % (len(prods), len(pairs), len(news)))

hits = []
for pool, label in ((prods, "products"), (pairs, "pairs"), (news, "news")):
    for tag in TAGS:
        i, p, _ = locate(pool, tag, label)
        if p is not None:
            hits.append((label, i, tag, p))
            break

for label, i, tag, p in hits:
    nm = (p.get("names") or [p.get("name", "")])[0]
    print("\n[%s idx=%d] %s" % (label, i, nm[:70]))
    print("   中文品名:", similar.cn_for(p))
    print("   inq:", p.get("inq"), " spus:", p.get("spus"))
    sig, _ = similar.sig_for_url(p.get("img"))
    print("   sig:", "有" if sig else "无",
          " shape=%d" % (sig[4] if sig else -1))
    print("   img:", (p.get("img") or "")[:78])

# 两两互比
print("\n===== 两两图距 =====")
for a in hits:
    for b in hits:
        if a[1] >= b[1] and a[0] == b[0]:
            continue
        if a[0] == b[0] and a[1] == b[1]:
            continue
        sa, _ = similar.sig_for_url(a[3].get("img"))
        sb_, _ = similar.sig_for_url(b[3].get("img"))
        if not sa or not sb_:
            continue
        print("\n%s(idx=%d)[%s]" % (a[2][:24], a[1], similar.cn_for(a[3])))
        print("  vs %s(idx=%d)[%s]" % (b[2][:24], b[1], similar.cn_for(b[3])))
        print("  h=%d s=%d p=%d fine=%d" % (
            similar.hdist(sa, sb_), similar.sdist(sa, sb_),
            similar.shapdist(sa, sb_), similar.fine_dist(sa, sb_)))
        break
