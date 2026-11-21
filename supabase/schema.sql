-- ============================================================
-- Mico 选品工作台 · Supabase 建表脚本
-- 使用方法：Supabase 控制台 → SQL Editor → 粘贴本文件 → Run
-- ============================================================

-- ① 已询盘产品库（主数据）
create table if not exists products (
  id          bigserial primary key,
  row_no      int,                          -- 原 Excel 行号
  names       text[] not null default '{}', -- 多语言标题（保留全部）
  url         text not null,
  url_key     text not null,                -- 规范化链接（去协议/参数/尾斜杠，小写）
  inq         text,                         -- 询盘号
  batch       text,                         -- 批次号（询盘号末4位）
  spu         text,                         -- SPU（多个用空格分隔）
  spus        text[] not null default '{}', -- SPU 数组
  u1688       text[] not null default '{}', -- 1688 链接
  img         text,                         -- 产品图 URL
  created_at  timestamptz default now(),
  updated_at  timestamptz default now()
);

create unique index if not exists products_url_key_idx on products (url_key);
create index if not exists products_spu_idx on products (spu);
create index if not exists products_inq_idx on products (inq);

-- ② 待采集配对表（每日订单中「已询盘但无 SKU」的产品）
create table if not exists pairs (
  id          bigserial primary key,
  work_date   date not null default current_date,
  url         text,
  names       text[] not null default '{}',
  inq         text,
  spu         text,                          -- 原有 SPU
  filled_spu  text,                          -- 人工填写的 SPU
  u1688       text[] not null default '{}',
  img         text,
  note        text,
  mark        text,                          -- same / diff（相似款标记）
  same_as     bigint,                        -- 标记为同款的产品 id
  edited_by   text,                          -- 最后编辑人
  created_at  timestamptz default now(),
  updated_at  timestamptz default now()
);

create index if not exists pairs_date_idx on pairs (work_date desc);
create index if not exists pairs_inq_idx on pairs (inq);

-- ③ 未询盘新品（待生成询盘需求表）
create table if not exists news (
  id          bigserial primary key,
  work_date   date not null default current_date,
  url         text,
  name        text,
  names       text[] not null default '{}',
  img         text,
  skus        text[] not null default '{}',
  l1          text,                          -- 一级分类
  l2          text,                          -- 二级分类
  l3          text,                          -- 三级分类
  checked     boolean default true,
  edited_by   text,
  created_at  timestamptz default now(),
  updated_at  timestamptz default now()
);

create index if not exists news_date_idx on news (work_date desc);

-- ④ 编辑日志（谁改了什么）
create table if not exists edit_log (
  id          bigserial primary key,
  ts          timestamptz default now(),
  editor      text,
  table_name  text,
  target_id   bigint,
  old_spu     text,
  new_spu     text,
  note        text
);

create index if not exists edit_log_ts_idx on edit_log (ts desc);

-- ⑤ 操作日志（上传文件等）
create table if not exists op_log (
  id          bigserial primary key,
  ts          timestamptz default now(),
  editor      text,
  action      text,
  detail      jsonb
);

-- ============================================================
-- 行级安全（RLS）
-- 使用 anon key 时默认拒绝，这里放开给已登录团队使用
-- 若要更严格，可改用 Supabase Auth 并校验 auth.uid()
-- ============================================================
alter table products  enable row level security;
alter table pairs     enable row level security;
alter table news      enable row level security;
alter table edit_log  enable row level security;
alter table op_log    enable row level security;

do $$
declare t text;
begin
  foreach t in array array['products','pairs','news','edit_log','op_log']
  loop
    execute format('drop policy if exists %I_all on %I;', t || '_rw', t);
    execute format(
      'create policy %I on %I for all using (true) with check (true);',
      t || '_rw', t
    );
  end loop;
end $$;

-- ============================================================
-- 触发器：自动更新 updated_at
-- ============================================================
create or replace function touch_updated_at() returns trigger as $$
begin
  new.updated_at = now();
  return new;
end $$ language plpgsql;

do $$
declare t text;
begin
  foreach t in array array['products','pairs','news']
  loop
    execute format('drop trigger if exists %I on %I;', t || '_touch', t);
    execute format(
      'create trigger %I before update on %I for each row execute function touch_updated_at();',
      t || '_touch', t
    );
  end loop;
end $$;

-- ============================================================
-- 验证
-- ============================================================
select 'products' as tbl, count(*) from products
union all select 'pairs', count(*) from pairs
union all select 'news', count(*) from news;
