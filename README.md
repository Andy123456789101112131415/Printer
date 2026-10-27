# 打印站

一个简单的局域网/公网打印网站：每个人选**房间号**、填**姓名**、上传文件，管理员按房间查看并**一键打印整个房间**的所有文件。

## 功能

- 上传页：房间号 + 姓名 + 多文件上传
- 管理后台（密码保护）：按房间分组查看、下载、删除、预览
- 按房间打印，两种方式：
  1. **打印视图（浏览器打印）**：把整个房间的文件排版成一页，按 `Ctrl+P` 即可打印，任何文件类型都能用
  2. **服务器直接打印**：后台静默调用本机打印机批量打印

## 环境要求

- Windows + Python 3.9+
- Microsoft Office（Word/Excel/PowerPoint，用于 Office 文件的转换与打印）
- 依赖：`pip install -r requirements.txt`（Flask、waitress 生产服务器、pywin32）

## 启动

```powershell
cd C:\Users\1\Desktop\打印
pip install -r requirements.txt
python app.py
```

打开 http://127.0.0.1:5000 上传文件，管理员入口 http://127.0.0.1:5000/admin

- 管理密码：存在 `.admin_password` 文件里（当前为 `AndySmart123`）
- 修改密码：直接编辑 `.admin_password` 文件，或设置环境变量后重启（环境变量优先）
  ```powershell
  $env:ADMIN_PASSWORD = "你的新密码"
  python app.py
  ```

## 打印设置

### 方式一：打印视图（推荐，立即可用）

管理员进入房间页 → 点「打印视图（浏览器打印）」→ 按 `Ctrl+P`。
Office 文件会在后台自动转成 PDF 再嵌入页面。

### 方式二：服务器直接打印

房间页选择打印机（留空 = 默认打印机）→ 点「服务器直接打印整间房」。

- Word/Excel/PowerPoint/TXT/CSV：直接通过 Office 打印，无需额外配置
- 图片：通过 Windows 系统打印
- PDF：按优先级自动选用
  1. **SumatraPDF**（已内置于 `tools\SumatraPDF.exe`）：PDF 打印最稳，且能强制 A4（`paper=A4`）
  2. Adobe Reader（如已安装）
  3. Edge 静默打印（兜底）

**打印机记忆**：房间页打印机下拉框会自动选中上次记住的打印机；点「记住当前打印机为默认」可固定，下次进任意房间都自动选中。

**纸张与任务名**：

- **全部按 A4 打印**：Word/Excel 在打印前强制设置 A4；PDF 用 SumatraPDF 以 `paper=A4` 打印；PPT 先转 PDF 再按 A4 打印。
- **任务名显示为「房间号_姓名」**：打印前把文件复制成同名临时副本再打印，所以打印机队列/屏幕上看到的是 `301_张三.docx`，方便取件时辨认。同房间同姓名多个文件会自动加序号（`301_张三_2.docx`）。网页上仍显示原始文件名。
- 上传的文件若本身是 A3 版式，也会被强制按 A4 输出。

### 学校教师卡打印（刷卡取件）

学校的打印机多为「漫游打印/刷卡取件」：电脑打印后，到打印机刷教师卡才会出纸。

**本机当前配置（已接好）**：

| 项 | 值 |
|----|-----|
| 打印机型号 | Canon imageRUNNER ADVANCE C5840 |
| 打印机 IP | `172.16.14.98` |
| Windows 打印机名 | `\\172.16.14.98\wccprinter` |
| 驱动 | Canon Generic Plus UFR II |
| 队列名（IPP） | `WCIS_AirPrinterA` |

已设为站点默认打印机，进任意房间页会自动选中。

**打印流程**：

1. 房间页确认打印机是 `\\172.16.14.98\wccprinter` → 点「服务器直接打印整间房」
2. 提示「已提交，请到打印机刷教师卡取件」
3. 拿教师卡到打印机刷一下，整批文件一起出纸

**换一台电脑接入时**：在 Windows「设置 → 蓝牙和其他设备 → 打印机和扫描仪 → 添加设备」里按 IP 添加，或运行：

```powershell
Add-Printer -ConnectionName "\\172.16.14.98\wccprinter"
```

然后在房间页选中它并点「记住当前打印机为默认」。

> 若房间页出现「默认打印机当前不可用」警告，说明没连上校园网或打印机离线。

## 公网访问（cpolar）

已安装 cpolar（`%LOCALAPPDATA%\cpolar\cpolar\cpolar.exe`，无需管理员权限），并配置了**固定公网地址**：

**https://print123.cpolar.cn**

**启动方式（推荐）**：双击 [启动打印站.bat](启动打印站.bat)，会自动启动网站 + 固定地址隧道。

> 固定地址由 cpolar 专业版的「保留子域名 print123」提供，断线重连地址也不会变。

**查看当前公网地址**（备用）：

```powershell
python show_url.py
```

> 管理密码已设置（见上文），公网开放前请确认已修改默认密码。

**手动启动**（不用 bat 时）：

```powershell
python app.py                       # 窗口 1：网站
& "$env:LOCALAPPDATA\cpolar\cpolar\cpolar.exe" http 5000 -log stdout -region cn -subdomain print123   # 窗口 2：隧道
```

## 数据存储

- 上传文件：`uploads\`（文件名随机化，原始文件名存数据库）
- 文件元数据：`data.db`（SQLite）
- 预览缓存：`uploads\.preview\`

备份时整目录拷贝即可。

## Git 版本管理

项目已关联 GitHub 仓库 `Andy123456789101112131415/Printer`（SSH 推送）。

**自动保存与推送**：

```powershell
python autosave.py
```

脚本会监控项目文件，修改后静默 60 秒（可传参修改，如 `python autosave.py 30`）自动 commit 并 push。
`uploads\`、`data.db`、`.secret_key`、`.admin_password` 等已在 `.gitignore` 中排除，不会被提交。

## 安全说明

- 管理后台有密码 + 会话登录 + CSRF 防护
- 上传限制：单次请求最大 100MB，白名单扩展名
- 房间号过滤非法字符，存储文件名随机化防止路径穿越
