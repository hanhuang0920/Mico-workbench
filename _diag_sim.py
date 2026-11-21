# -*- coding: utf-8 -*-
"""诊断：这两组同款产品为何没比出来（临时脚本）"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import supabase_store as sb
import similar

TARGETS = [
    "Lederreparaturgel für Sitze",
    "Zaawansowany Żel Naprawczy",
    "Szelki dla Psa Regulowane",
    "Hundegeschirr gepolstert",
]

prods = sb.get_products()
pairs = sb.get_pairs()
news = sb.get_news()
pool = prods + pairs + news

found = {}
for p in pool:
    nm = (p.get("names") or [p.get("name", "")])[0]
    for t in TARGETS:
        if t.lower() in nm.lower() and t not in found:
            found[t] = p

for t in TARGETS:
    p = found.get(t)
    if not p:
        print("未找到:", t)
        continue
    nm = (p.get("names") or [""])[0]
    cn = similar.cn_for(p)
    sig, _ = similar.sig_for_url(p.get("img"))
    print("\n=== %s" % t)
    print("   标题:", nm[:80])
    print("   中文品名:", repr(cn))
    print("   img:", (p.get("img") or "")[:80])
    print("   sig:", "有" if sig else "无")

# 两两比对距离
pairs_to_cmp = [
    ("Lederreparaturgel für Sitze", "Zaawansowany Żel Naprawczy"),
    ("Szelki dla Psa Regulowane", "Hundegeschirr gepolstert"),
]
print("\n===== 图距明细 =====")
for a, b in pairs_to_cmp:
    pa, pb = found.get(a), found.get(b)
    if not pa or not pb:
        continue
    sa, _ = similar.sig_for_url(pa.get("img"))
    sb_, _ = similar.sig_for_url(pb.get("img"))
    if not sa or not sb_:
        print("%s vs %s: 缺签名" % (a, b))
        continue
    na = (pa.get("names") or [""])[0]
    nb = (pb.get("names") or [""])[0]
    d = similar.hdist(sa, sb_)
    ds = similar.sdist(sa, sb_)
    dp = similar.shapdist(sa, sb_)
    df = similar.fine_dist(sa, sb_)
    cn_a, cn_b = similar.cn_for(pa), similar.cn_for(pb)
    ns = max(similar.nsim(na, n) for n in (pb.get("names") or [nb]))
    print("\n%s  [%s]" % (a, cn_a))
    print("  vs %s  [%s]" % (b, cn_b))
    print("  hdist=%d  sdist=%d  shapdist=%d  fine=%d  标题相似=%.2f" % (d, ds, dp, df, ns))
    print("  tiny判定(h<=40或s<=55或p<=45):", (d <= 40 or ds <= 55 or dp <= 45))
    print("  中文品名相同:", cn_a == cn_b)
