# -*- coding: utf-8 -*-
"""一次性预计算：下载全部产品图 → 算签名 → 写入 Supabase img_sigs 表。
本地跑（网速快、结果持久），之后云端比对不用再下载 365 张图。

用法：
  $env:SUPABASE_URL="..."
  $env:SUPABASE_ANON_KEY="..."
  python build_sig_cache.py
"""
import os
import sys
import time
import threading

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import supabase_store as sb
import similar

if not sb.enabled():
    print("请先设置 SUPABASE_URL / SUPABASE_ANON_KEY 环境变量")
    sys.exit(1)


def main():
    print("[1/4] 读取产品与已有签名…")
    prods = sb.get_products()
    pairs = sb.get_pairs()
    urls = []
    seen = set()
    for p in prods + pairs:
        u = (p.get("img") or "").strip()
        if u.startswith("http"):
            k = similar.norm(u)
            if k not in seen:
                seen.add(k)
                urls.append(u)
    have = sb.get_img_sigs()
    todo = [u for u in urls if similar.norm(u) not in have]
    print("  图片总数 %d，已有签名 %d，待计算 %d" % (len(urls), len(have), len(todo)))
    if not todo:
        print("✅ 全部已有签名，无需计算")
        return

    print("[2/4] 并发下载并计算签名…")
    results = {}
    failed = []
    lock = threading.Lock()
    done = [0]

    def work(u):
        p = similar.ensure_img(u)
        sig = similar.img_sig(p) if p else None
        with lock:
            done[0] += 1
            if sig:
                results[similar.norm(u)] = [x for x in sig]
            else:
                failed.append(u)
            if done[0] % 50 == 0:
                print("  进度 %d/%d（失败 %d）" % (done[0], len(todo), len(failed)))

    threads = []
    queue = list(todo)

    def runner():
        while True:
            if not queue:
                break
            u = queue.pop(0)
            work(u)

    for _ in range(8):
        t = threading.Thread(target=runner)
        t.start()
        threads.append(t)
    for t in threads:
        t.join()

    print("  计算完成：成功 %d，失败 %d" % (len(results), len(failed)))

    print("[3/4] 批量写入 img_sigs…")
    n = sb.put_img_sigs_bulk(list(results.items()))
    print("  已写入 %d 条" % n)

    print("[4/4] 校验…")
    chk = sb.get_img_sigs()
    print("  img_sigs 现有 %d 条签名" % len(chk))
    if failed:
        print("  失败图片（下次比对时会自动补算）：")
        for u in failed[:10]:
            print("   ", u[:90])
    print("✅ 完成。相似款比对将直接读签名，无需下载图片")


if __name__ == "__main__":
    t0 = time.time()
    main()
    print("总耗时 %.1f 秒" % (time.time() - t0))
