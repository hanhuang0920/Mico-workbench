# -*- coding: utf-8 -*-
"""检查签名原始值，区分「真相同」与「计算失败」"""
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


for kw in ("Lederreparaturgel für Sitze", "Zaawansowany Żel Naprawczy",
           "Szelki dla Psa Regulowane", "Hundegeschirr gepolstert",
           "Wintersocken mit Schneemuster"):
    p = find(kw)
    if not p:
        print("未找到:", kw)
        continue
    sig, path = similar.sig_for_url(p.get("img"))
    if not sig:
        print("%-28s 无签名" % kw[:28])
        continue
    # sig = (ahash, dhash, edge, sat, shape, f1, f2, white, hue)
    zeros = [i for i, v in enumerate(sig[:7]) if not v]
    print("%-28s ahash=%d dhash=%d edge=%d sat=%d shape=%d f1=%d f2=%d white=%d 零值位=%s"
          % (kw[:28], sig[0], sig[1], sig[2], sig[3], sig[4], sig[5], sig[6], sig[7], zeros))
