-- ============================================================
-- 多客户改造：给三张主表 + 图片签名表加 customer 字段
-- 在 Supabase 控制台 → SQL Editor 执行本文件
-- 时间：2026-11-27
-- 执行后：现有数据全部归到 Mico，Stijn 从空白开始
-- ============================================================

-- 1) 产品库
alter table products add column if not exists customer text not null default 'Mico';
update products set customer = 'Mico' where customer is null or customer = '';

-- 2) 待采集配对
alter table pairs add column if not exists customer text not null default 'Mico';
update pairs set customer = 'Mico' where customer is null or customer = '';

-- 3) 未询盘新品
alter table news add column if not exists customer text not null default 'Mico';
update news set customer = 'Mico' where customer is null or customer = '';

-- 4) 图片签名缓存：按客户隔离（同 URL 不同客户各存一份）
alter table img_sigs add column if not exists customer text not null default 'Mico';
update img_sigs set customer = 'Mico' where customer is null or customer = '';

-- 5) 索引：按客户过滤要快
create index if not exists products_customer_idx on products (customer);
create index if not exists pairs_customer_idx on pairs (customer);
create index if not exists news_customer_idx on news (customer);
create index if not exists img_sigs_customer_idx on img_sigs (customer);

-- 6) 产品库唯一约束要带上客户：不同客户可以有同一个链接
drop index if exists products_url_inq_uniq;
create unique index if not exists products_customer_url_inq_uniq
  on products (customer, url_key, coalesce(inq, ''));

-- 7) 图片签名唯一约束同样带上客户
drop index if exists img_sigs_url_key_idx;
create unique index if not exists img_sigs_customer_url_uniq
  on img_sigs (customer, url_key);

-- 8) 校验
select 'products' as t, customer, count(*) from products group by customer
union all
select 'pairs', customer, count(*) from pairs group by customer
union all
select 'news', customer, count(*) from news group by customer;
