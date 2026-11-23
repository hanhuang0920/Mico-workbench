-- ============================================================
-- 在 Supabase 控制台 → SQL Editor 执行本文件
-- 目的：允许同一链接对应多个询盘号（同款被询盘两次 / 多站点各自被询盘）
-- 时间：2026-11-23
-- ============================================================

-- 1) 先记下当前约束（便于回滚）
--    原约束名：products_url_key_idx  (unique on products(url_key))

-- 2) 解除 url_key 唯一约束，改为普通索引（加速查询）
drop index if exists products_url_key_idx;
create index if not exists products_url_idx on products (url_key);

-- 3) 新增「同链接 + 同询盘号」唯一约束：真正的重复仍会被拦住
create unique index if not exists products_url_inq_uniq
  on products (url_key, coalesce(inq, ''));

-- 4) 校验：确认索引状态
select indexname, indexdef
from pg_indexes
where tablename = 'products'
order by indexname;
