# -*- coding: utf-8 -*-
"""探针：逐项打印案例1的判定中间值"""
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


a = find("Lederreparaturgel für Sitze")
b = find("Zaawansowany Żel Naprawczy")
print("target:", (a.get("names") or [""])[0][:60])
print("cand  :", (b.get("names") or [""])[0][:60])

tsig, _ = similar.sig_for_url(a.get("img"))
xsig, _ = similar.sig_for_url(b.get("img"))
d = similar.hdist(tsig, xsig)
ds = similar.sdist(tsig, xsig)
dp = similar.shapdist(tsig, xsig)
df = similar.fine_dist(tsig, xsig)
thue = similar._hue_of(tsig)
xhue = similar._hue_of(xsig)
hc = similar.cosv(thue, xhue) if xsig else 0.0
twhite = similar._white_of(tsig) if tsig else 0.0
xwhite = similar._white_of(xsig) if xsig else 0.0
white_pair = twhite > 0.45 and xwhite > 0.45
ov, ca, cb = similar.cat_overlap(a.get("names") or [], b.get("names") or [])
ns = max((similar.nsim((a.get("names") or [""])[0], n) for n in (b.get("names") or [])), default=0.0)

print("d=%d ds=%d dp=%d df=%d" % (d, ds, dp, df))
print("hc=%.3f  twhite=%.2f xwhite=%.2f white_pair=%s" % (hc, twhite, xwhite, white_pair))
print("ns=%.2f  ov=%s (ca=%s cb=%s)" % (ns, bool(ov), ca, cb))

has_img = xsig is not None and tsig is not None
d_eff = min(d, ds, dp, df // 6) if has_img else 9999
fine_score = max(0.0, (1200 - min(df, 1200)) / 1200.0)
small_score = max(0.0, (300 - min(d_eff, 300)) / 300.0)
score = (0.40 * fine_score + 0.25 * hc + 0.20 * small_score
         + 0.10 * ns + (0.05 if ov else 0.0))
print("d_eff=%d fine_score=%.3f small_score=%.3f score=%.3f" % (d_eff, fine_score, small_score, score))

tiny = has_img and (d <= 40 or ds <= 55 or dp <= 45)
print("tiny=%s" % tiny)
if tiny:
    lvl = "same"
elif has_img and df <= 600 and hc >= 0.75 and not white_pair:
    lvl = "same"
elif has_img and df <= 600 and ov and ns >= 0.30:
    lvl = "same"
elif has_img and df <= 700 and hc >= 0.55 and not white_pair:
    lvl = "likely"
elif has_img and df <= 750 and ov and ns >= 0.25:
    lvl = "likely"
elif has_img and df <= 800 and hc >= 0.50 and not white_pair:
    lvl = "possible"
elif has_img and df <= 850 and ov and ns >= 0.20:
    lvl = "possible"
elif ns >= 0.78:
    lvl = "likely"
elif ns >= 0.65 and ov:
    lvl = "possible"
else:
    lvl = "DROP"
print("level=%s" % lvl)
print("min_score=0.50 通过:", (score >= 0.50 or lvl in ("same", "likely")) and not (lvl == "possible" and score < 0.50))
