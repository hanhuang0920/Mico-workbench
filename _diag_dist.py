# -*- coding: utf-8 -*-
"""量化：目标 vs 同款 vs 误报图的距离分布，确定安全阈值"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import supabase_store as sb
import similar

pool = sb.get_products() + sb.get_pairs() + sb.get_news()


def find(kw):
    for p in pool:
        nm = (p.get("names") or [p.get("name", "")])[0]
        if kw.lower() in nm.lower():
            return p
    return None


def dist(a, b):
    sa, _ = similar.sig_for_url(a.get("img"))
    sb_, _ = similar.sig_for_url(b.get("img"))
    if not sa or not sb_:
        return None
    return (similar.hdist(sa, sb_), similar.sdist(sa, sb_),
            similar.shapdist(sa, sb_), similar.fine_dist(sa, sb_))


target = find("Lederreparaturgel für Sitze")
same = find("Zaawansowany Żel Naprawczy")
noise = [find("Wintersocken mit Schneemuster"), find("Shawl Wrap for Dress"),
         find("Kaschmir-Cardigan für Damen")]

print("目标:", (target.get("names") or [""])[0][:50])
print("\n【同款】")
d = dist(target, same)
if d:
    print("  %-46s h=%d s=%d p=%d fine=%d" % ((same.get("names") or [""])[0][:46], d[0], d[1], d[2], d[3]))
else:
    print("  缺签名")
print("\n【误报】")
for n in noise:
    if not n:
        continue
    d = dist(target, n)
    if d:
        print("  %-46s h=%d s=%d p=%d fine=%d" % ((n.get("names") or [""])[0][:46], d[0], d[1], d[2], d[3]))

# 全库统计：与目标的距离分布，看同款是否落在可区分区间
print("\n【全库最近 8 个】")
rows = []
tsig, _ = similar.sig_for_url(target.get("img"))
for p in pool:
    if p is target:
        continue
    sig, _ = similar.sig_for_url(p.get("img"))
    if not sig or not tsig:
        continue
    rows.append((similar.fine_dist(tsig, sig), similar.hdist(tsig, sig),
                 similar.sdist(tsig, sig), similar.shapdist(tsig, sig),
                 (p.get("names") or [""])[0], similar.cn_for(p)))
rows.sort()
for f, h, s, sh, nm, cn in rows[:8]:
    print("  fine=%4d h=%3d s=%3d p=%3d  [%s] %s" % (f, h, s, sh, cn, nm[:42]))
