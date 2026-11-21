# Mico 选品工作台 · 从 GitHub 开始的部署清单（黄翰版）

> 代码已整理好：`D:\DSH\mico-workbench`（全部通过本地启动 + Supabase 连通性验证）。
> 你只需要：装 Git → 建 GitHub 仓库 → 推代码 → Render 建服务 → 打开网址。

---

## 0. 本机状态（已替你确认）

| 项目 | 状态 |
|---|---|
| Git | ❌ 未安装（PATH 里没有；`C:\Program Files\Git` 也没有） |
| Python | ✅ `C:\Users\Mayn\AppData\Local\Python\bin\python.exe`（3.14.4，openpyxl/Pillow/numpy 已装好） |
| GitHub CLI | ❌ 未安装（可选，装了可以少点几次鼠标） |
| 项目代码 | ✅ `D:\DSH\mico-workbench`（含 `app.py`、`similar.py`、`supabase_store.py`、`_imgmap.json`、`已询盘产品_四列表_精简版.xlsx`） |
| Supabase | ✅ 库已可用，5 张表都在：`products` 365 条、`pairs` 20 条、`news` 16 条、`edit_log` 0、`op_log` 0 |
| 本地自测 | ✅ `python app.py` 能起，登录 → 首页渲染 365/20/16 条数据，图片链接正常 |

> 注意：交接说明里写的 `已询盘产品_四列表.xlsx`（约 2MB，含内嵌图）不需要。
> `app.py` 优先读 `已询盘产品_四列表_精简版.xlsx`（55KB，无内嵌图），云上省空间、启动更快。

---

## 1. 安装 Git（必做）

1. 打开 https://git-scm.com/download/win
2. 下载 **64-bit Git for Windows Setup**，一路默认「Next」装完
3. **关掉所有终端/编辑器再重开**（PATH 才会刷新）

验证：新开 PowerShell 输入 `git --version`，出现版本号即可。

> 可选：再装 GitHub CLI https://cli.github.com/ ，后面第 3 步可以直接 `gh repo create` 一键建库。

---

## 2. 在 GitHub 建仓库

网页操作（最稳）：

1. 登录 https://github.com
2. 右上角 **+** → **New repository**
3. 填：
   - **Repository name**：`mico-workbench`
   - **Visibility**：**Private**（私有，代码里带 Supabase 地址，别公开）
   - **不要**勾 Add a README / .gitignore / license（本地已有）
4. **Create repository**，记下仓库地址，形如
   `https://github.com/你的账号/mico-workbench.git`

---

## 3. 推送代码到 GitHub

打开 PowerShell，逐条执行：

```powershell
cd D:\DSH\mico-workbench

git init
git branch -M main

# 第一次用 Git 需要设一次身份（换成你自己的）
git config user.name "huanghan"
git config user.email "你的GitHub邮箱"

git add .
git commit -m "Mico workbench: Supabase + Render"

git remote add origin https://github.com/你的账号/mico-workbench.git
git push -u origin main
```

`git push` 时会弹浏览器让你登录 GitHub 授权（Git Credential Manager 自动处理），点同意即可。

**预期**：仓库里出现 `app.py`、`similar.py`、`supabase_store.py`、`requirements.txt`、`Procfile`、`runtime.txt`、`_imgmap.json`、`已询盘产品_四列表_精简版.xlsx`。
**不应出现**：`_imgcache/`、`备份/`、`_data.json`、`_sigs.json`（`.gitignore` 已排除）。

---

## 4. Render 建 Web Service

1. https://render.com → **Sign in with GitHub**
2. 右上 **New** → **Web Service**
3. 选 `mico-workbench` 仓库 → **Connect**
4. 填写：

| 配置项 | 填写 |
|---|---|
| Name | `mico-workbench` |
| Language / Runtime | `Python 3` |
| Branch | `main` |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `python app.py` |
| Instance Type | **Free** |

5. **Environment Variables**（一条一条 Add）：

```
HOST              = 0.0.0.0
APP_PASSWORD      = mico2026
APP_USERS         = 李锦泳,黄翰
SUPABASE_URL      = https://ollosmzqgmsivpruwfua.supabase.co
SUPABASE_ANON_KEY = sb_publishable_CxD-zwVGCgRY49jYHrQWqA_8OhkR5J8
```

> `HOST=0.0.0.0` 必须加，否则 Render 从外部访问不到（代码默认监听 127.0.0.1）。
> `PORT` 不用设，Render 自动注入。
> 数据全在 Supabase，Render 不需要挂磁盘卷。

6. **Create Web Service** → 看 Logs 等 3–5 分钟，出现 `监听: http://0.0.0.0:xxxx` 即成功。

---

## 5. 验证与使用

打开 Render 给的地址 `https://mico-workbench.onrender.com`：

1. 选「我是谁」→ 李锦泳 / 黄翰
2. 输入密码 `mico2026`
3. 三个页面：**产品库** / **待采集配对** / **未询盘新品**

首次打开慢约 30 秒（免费层休眠唤醒），属正常。

---

## 6. 两个已知小坑

- **云端「写回四列表」不保留**：`保存 SPU 并写回四列表` 会同时做两件事——更新**线上 xlsx 文件**（`wb.save(INQ_FILE)`，`app.py:551`）和写 **Supabase `products`**。
  Render 免费层重启/重新部署后文件系统还原，xlsx 上的改动会丢；**Supabase 里的改动不会丢**，重启后读的仍是 Supabase。
  所以云上放心点它；需要拿到最新 Excel 时，在本地跑 `python app.py` 再点一次即可。
- **中文文件名**：`已询盘产品_四列表_精简版.xlsx` 是中文名，Git 和 Render 都能处理，但如果 build 报找不到文件，把本地文件名改成 `data.xlsx`，同时把 `app.py` 第 38 行改成 `data.xlsx` 再推一次。

---

## 7. 以后怎么更新

```powershell
cd D:\DSH\mico-workbench
git add .
git commit -m "改了什么"
git push
```

Render 检测到 `main` 有新提交会自动重新部署。
