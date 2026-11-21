# -*- coding: utf-8 -*-
"""端到端验证：这两组同款能否互相比对出来（临时脚本）"""
import sys
import os
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import supabase_store as sb
import similar

CASES = [
    ("Lederreparaturgel für Sitze", "Zaawansowany Żel Naprawczy"),
    ("Szelki dla Psa Regulowane", "Hundegeschirr gepolstert"),
]

pool = sb.get_products() + sb.get_pairs() + sb.get_news()


def find(keyword):
    for p in pool:
        nm = (p.get("names") or [p.get("name", "")])[0]
        if keyword.lower() in nm.lower():
            return p
    return None


for a_kw, b_kw in CASES:
    a, b = find(a_kw), find(b_kw)
    print("\n" + "=" * 60)
    if not a or not b:
        print("未找到:", a_kw if not a else b_kw)
        continue
    print("目标:", (a.get("names") or [""])[0][:56], "|", similar.cn_for(a))
    print("期望命中:", (b.get("names") or [""])[0][:56], "|", similar.cn_for(b))
    t = time.time()
    res = similar.find_similar(a.get("names") or [], a.get("img") or "", pool,
                               exclude_inq=a.get("inq") or None, t_url=a.get("url"))
    dt = time.time() - t
    print("比对 %.2fs，命中 %d 个：" % (dt, len(res)))
    hit = False
    for r in res:
        mark = ""
        if (r.get("name") or "").lower() == (b.get("names") or [""])[0].lower():
            mark = "  ← 目标同款"
            hit = True
        print("   %.3f %-10s %s%s" % (r.get("score"), r.get("level"),
                                      (r.get("name") or "")[:52], mark))
    print("结论:", "✅ 能找到对方" if hit else "❌ 仍未找到")
