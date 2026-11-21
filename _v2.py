# -*- coding: utf-8 -*-
"""线上验证这两组同款能否比出来"""
import http.client
import ssl
import json
import re

ctx = ssl.create_default_context()


def raw(method, path, body=None, headers=None, timeout=60):
    c = http.client.HTTPSConnection("mico-workbench.onrender.com", 443,
                                    timeout=timeout, context=ctx)
    h = {"User-Agent": "mico-verify", "Connection": "close"}
    if headers:
        h.update(headers)
    c.request(method, path, body=body, headers=h)
    r = c.getresponse()
    d = r.read()
    hd = r.getheaders()
    c.close()
    return r.status, hd, d


s, hd, b = raw("POST", "/login", "user=%E9%BB%84%E7%BF%B0&pwd=mico2026",
               {"Content-Type": "application/x-www-form-urlencoded"})
print("login:", s)
ck = "; ".join(v.split(";")[0] for k, v in hd if k.lower() == "set-cookie")
s2, _, b2 = raw("GET", "/", headers={"Cookie": ck})
html = b2.decode("utf-8", "replace")
print("home:", s2, "产品卡:", html.count('<div class="pc"'))

KEYS = ["Lederreparaturgel", "Zaawansowany", "Szelki dla Psa", "Hundegeschirr",
        "Krem do Naprawy"]

for kw in KEYS:
    # 在 paircard 或 newscard 里找
    idx = None
    kind = None
    for m in re.finditer(r'<div class="(paircard|newscard)"[^>]*data-idx="(\d+)"(.*?)(?=<div class="(?:paircard|newscard)"|<div class="panel" id="p2)', html, re.S):
        if kw.lower() in m.group(3).lower():
            idx = int(m.group(2))
            kind = "pair" if m.group(1) == "paircard" else "news"
            break
    if idx is None:
        print("%-22s 不在配对/新品卡片里" % kw)
        continue
    s3, _, b3 = raw("POST", "/api/similar", json.dumps({"idx": idx, "kind": kind}),
                    {"Content-Type": "application/json", "Cookie": ck})
    d = json.loads(b3.decode())
    cs = d.get("candidates") or []
    tcn = (d.get("target") or {}).get("cn")
    print("\n[%s idx=%d] %s → 命中 %d" % (kind, idx, tcn, len(cs)))
    for c in cs:
        print("    %.3f %-8s [%s] %s" % (c.get("score"), c.get("level"), c.get("cn"),
                                         (c.get("name") or "")[:46]))
