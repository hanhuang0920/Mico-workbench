# -*- coding: utf-8 -*-
"""
把本地 Excel + JSON 数据迁移到 Supabase

用法（先设好环境变量）：
  set SUPABASE_URL=https://xxx.supabase.co
  set SUPABASE_ANON_KEY=eyJhbGci...
  python migrate_to_supabase.py
"""
import os
import sys
import io
import json
import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

import supabase_store as sb


def load_local_products():
    """从已询盘四列表读取产品。"""
    import openpyxl
    f = os.path.join(BASE, "已询盘产品_四列表.xlsx")
    wi = openpyxl.load_workbook(f).active
    imgmap = {}
    mp = os.path.join(BASE, "_imgmap.json")
    if os.path.exists(mp):
        imgmap = json.load(open(mp, encoding="utf-8"))

    out = []
    for r in range(4, wi.max_row + 1):
        nm = wi.cell(row=r, column=2).value
        if not nm:
            continue
        names = [x for x in str(nm).split("\n") if x.strip()]
        url = str(wi.cell(row=r, column=3).value or "")
        iq = str(wi.cell(row=r, column=4).value or "").strip()
        spu_raw = str(wi.cell(row=r, column=5).value or "").strip()
        u1688 = [x.strip() for x in str(wi.cell(row=r, column=6).value or "").split("\n")
                 if x.strip().startswith("http")]
        import re
        m = re.search(r"INQ-(\d+)", iq)
        out.append({
            "row": r,
            "names": names,
            "url": url,
            "inq": iq,
            "batch": m.group(1)[-4:] if m else "",
            "spu": spu_raw,
            "spus": sb.split_spu(spu_raw),
            "u1688": u1688,
            "img": imgmap.get(sb.norm_url(url), ""),
        })
    return out


def main():
    if not sb.enabled():
        print("❌ 未配置环境变量 SUPABASE_URL / SUPABASE_ANON_KEY")
        print("   请先在 Supabase 控制台 Settings → API 获取后设置")
        return

    print("=" * 60)
    print("迁移到 Supabase")
    print("  URL: %s" % sb.SUPABASE_URL)
    print("=" * 60)

    # 1) 测试连接
    print("\n[1/4] 测试连接…")
    try:
        n = sb.count("products")
        print("  ✓ 连接成功（products 现有 %d 条）" % n)
    except Exception as e:
        print("  ✗ 连接失败: %s" % e)
        print("  请确认已执行 supabase/schema.sql 建表")
        return

    # 2) 产品
    print("\n[2/4] 导入产品库…")
    prods = load_local_products()
    print("  本地 %d 款" % len(prods))
    try:
        n = sb.upsert_products(prods)
        print("  ✓ 已写入 %d 款" % n)
    except Exception as e:
        print("  ✗ 失败: %s" % e)
        return

    # 3) 当日配对 / 新品
    print("\n[3/4] 导入当日配对与新品…")
    d = {}
    dp = os.path.join(BASE, "_data.json")
    if os.path.exists(dp):
        d = json.load(open(dp, encoding="utf-8"))

    today = datetime.date.today().isoformat()
    pairs = d.get("pairs", [])
    news = d.get("news", [])

    if pairs:
        try:
            sb.save_pairs(today, pairs)
            print("  ✓ 配对表 %d 条（日期 %s）" % (len(pairs), today))
        except Exception as e:
            print("  ✗ 配对表失败: %s" % e)
    else:
        print("  - 配对表无数据")

    if news:
        try:
            sb.save_news(today, news)
            print("  ✓ 新品 %d 条" % len(news))
        except Exception as e:
            print("  ✗ 新品失败: %s" % e)
    else:
        print("  - 新品无数据")

    # 4) 校验
    print("\n[4/4] 校验…")
    stats = sb.get_stats()
    for k, v in stats.items():
        print("  %-10s %d 条" % (k, v))

    sample = sb.get_products()[:3]
    print("\n  样例:")
    for p in sample:
        print("    R%-3s %-14s %s" % (p["row"], p["spu"] or "-", p["names"][0][:48]))

    print("\n✅ 迁移完成")


if __name__ == "__main__":
    main()
