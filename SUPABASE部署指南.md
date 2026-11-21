# Supabase 部署指南

> 全部使用 Supabase：数据库存数据、Storage 存图片、免费且不休眠。
> 后端仍需一个 Python 运行环境（用 Render 免费层即可）。

---

## 第一步：新建 Supabase 项目

1. 打开 https://supabase.com 登录
2. 点 **New project**
3. 填写：
   - **Name**：`mico-workbench`
   - **Database Password**：设一个强密码（记下来）
   - **Region**：选 **Southeast Asia (Singapore)** —— 国内访问最快
   - **Plan**：Free
4. 点 **Create new project**，等待约 2 分钟初始化

---

## 第二步：建表

1. 左侧菜单 → **SQL Editor**
2. 点 **New query**
3. 把 `supabase/schema.sql` 的全部内容粘贴进去
4. 点 **Run**（或 Ctrl+Enter）
5. 看到 `Success. No rows returned` 即成功

> 这一步会创建 5 张表：`products` `pairs` `news` `edit_log` `op_log`

---

## 第三步：获取密钥

左侧 **Project Settings**（齿轮图标）→ **API**

记下两个值：

| 名称 | 位置 | 说明 |
|---|---|---|
| **Project URL** | Project URL | 形如 `https://abcdefgh.supabase.co` |
| **anon public** | Project API keys → anon public | 形如 `eyJhbGciOi...`，很长一串 |

> ⚠️ 不要用 `service_role` 那个，它是管理员密钥，泄露有风险。

---

## 第四步：本地迁移数据

在 `D:\DSH\Mico` 目录开 PowerShell：

```powershell
$env:SUPABASE_URL="https://你的项目.supabase.co"
$env:SUPABASE_ANON_KEY="你的anon key"
python migrate_to_supabase.py
```

预期输出：

```
[1/4] 测试连接…
  ✓ 连接成功（products 现有 0 条）
[2/4] 导入产品库…
  本地 365 款
  ✓ 已写入 365 款
[3/4] 导入当日配对与新品…
  ✓ 配对表 20 条
  ✓ 新品 27 条
[4/4] 校验…
  products   365 条
  pairs      20 条
  news       27 条
✅ 迁移完成
```

---

## 第五步：部署后端到 Render（免费）

1. 打开 https://render.com 用 GitHub 登录
2. 把 `D:\DSH\Mico` 推到 GitHub 仓库（私有即可）
3. Render 点 **New** → **Web Service**
4. 连接你的仓库
5. 配置：

| 项 | 值 |
|---|---|
| **Name** | `mico-workbench` |
| **Runtime** | Python 3 |
| **Build Command** | `pip install -r requirements.txt` |
| **Start Command** | `python app.py` |
| **Instance Type** | Free |

6. 添加环境变量（**Environment** 页）：

```
HOST             = 0.0.0.0
APP_PASSWORD     = 你的团队密码
APP_USERS        = 李锦泳,同事B,同事C
SUPABASE_URL     = https://你的项目.supabase.co
SUPABASE_ANON_KEY= eyJhbGciOi...
```

7. 点 **Create Web Service**，等 3-5 分钟

部署完成后会给你一个网址：`https://mico-workbench.onrender.com`

---

## 第六步：同事使用

把网址 + 密码发给同事：

1. 打开网址
2. 选择自己的名字
3. 输入团队密码
4. 开始填 SPU

**数据自动同步**：所有人看到的是同一份数据，改完立刻保存到 Supabase。

---

## 免费额度

| 项目 | 免费额度 | 够用吗 |
|---|---|---|
| 数据库 | 500 MB | ✅ 365 款产品不到 5 MB |
| 带宽 | 5 GB/月 | ✅ 团队用足够 |
| 休眠 | **不休眠** | ✅ 比 Railway/Render 好 |
| Render 后端 | 15 分钟无访问休眠 | ⚠️ 首次打开慢约 30 秒 |

> Render 免费层会休眠，但每天早上打开一次就唤醒了。
> 如果介意，可以加一个定时 ping（如 UptimeRobot 免费版每小时访问一次）。

---

## 关于图片

产品图目前存的是 **Shopify CDN 链接**，不占 Supabase 空间。

如果以后要把图片也搬到 Supabase Storage：

1. 左侧 **Storage** → **New bucket** → 命名 `product-images` → 勾选 **Public**
2. 我另外写个脚本把 365 张图上传并替换链接

---

## 常见问题

**Q: 数据会丢吗？**
A: 不会。Supabase 免费层也有每日备份，且数据库持久化。

**Q: 两个人同时改同一个产品？**
A: 后保存的覆盖先保存的。建议分工：一人负责一半产品。
每次修改都会记入 `edit_log` 表，可查谁改的。

**Q: 想加/减成员？**
A: 改 Render 的 `APP_USERS` 环境变量，重启生效。

**Q: 改密码？**
A: 改 `APP_PASSWORD`，重启生效。已登录的人需要重新登录。

**Q: 怎么备份数据？**
A: Supabase 控制台 → Database → Backups，或自己写脚本导出 JSON。
