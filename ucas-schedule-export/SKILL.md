---
name: ucas-schedule-export
description: 抓取国科大（中国科学院大学）超星课表并导出为 WakeUp 课程表可导入的 Excel/CSV 文件。当用户提到 国科大/中科院大学/UCAS 课表、kb.mooc.ucas.edu.cn、学习通课表、课表爬取、导出课表、WakeUp 课程表导入、课表 Excel/CSV 时使用——即使用户没有明确说"导出"或"爬取"。涵盖登录态认证（Cookie 导入或学习通账号登录，两种方式可随时切换）、逐周爬取全部学期课程、按学期分文件导出。
---

# 国科大课表导出（WakeUp 格式）

帮用户通过 kb.mooc.ucas.edu.cn（超星课表国科大部署）的身份认证，爬取**所有学年学期**的课程，并按学期分别导出 WakeUp 课程表 App 可导入的 `.xlsx` / `.csv` 文件。

配套工具：本 skill 目录下的 `scripts/ucas_crawler.py`（下称 `CRAWLER`，用实际绝对路径替换）。所有登录态文件存放在 `~/.ucas_schedule/`，跨会话可复用。

## 工作流程

### 第 1 步：依赖检查

```bash
python3 <CRAWLER> doctor
```

缺什么按提示安装（`pip3 install openpyxl` / `pip3 install playwright && python3 -m playwright install chromium`）。Cookie 方式**不需要** playwright，只有浏览器登录需要。

### 第 2 步：认证方式（必须让用户二选一，不得跳过）

用 AskUserQuestion 让用户在两种认证方式中选择：

**方式 A：Cookie 导入**（适合：不想装 playwright、浏览器里已登录过）

用户选了这种方式后，**必须教用户怎么拿 Cookie**，逐条给出步骤：

1. 在自己的浏览器打开登录页并登录：
   `https://passport.mooc.ucas.edu.cn/login?fid=&newversion=true&refer=https%3A%2F%2Fkb.mooc.ucas.edu.cn%2Fres%2Fpc%2Fcurriculum%2Fschedule.html`
   （扫码/手机号+学习通密码/机构账号"中国科学院大学"+学号/短信验证码均可）
2. 登录跳回课表页后，按 F12 打开开发者工具 → **Network（网络）** 标签 → 刷新页面
3. 在请求列表里点任意一个发往 `kb.mooc.ucas.edu.cn` 的请求（如 `getMyLessons`）
4. 在 **Request Headers（请求标头）** 里找到 `Cookie:` 一行，**右键复制整串值**
5. 把整串发给模型

**关键警告（必须告知用户）**：一定要从 Network 的请求头里复制，**不要**用控制台敲 `document.cookie`——登录凭据包含 HttpOnly Cookie（如 vc3），JS 读不到，缺了它会被服务器 302 踢回登录页。

拿到后：

```bash
python3 <CRAWLER> save-cookie --cookie "用户复制的整串"
```

脚本会自动验证。若提示无效：多半是复制不完整或已过期，让用户重新复制，或建议切换到方式 B。

**方式 B：学习通账号登录**（适合：想要全自动、一次登录长期复用）

```bash
python3 <CRAWLER> login-browser
```

脚本会弹出一个 Chromium 窗口并打开登录页。**明确告诉用户**：
- 登录页地址是 `https://passport.mooc.ucas.edu.cn/login?fid=&newversion=true&refer=...`，在弹出的浏览器里操作即可，推荐**学习通 App 扫码**（不触发验证码）；也可手机号+学习通密码、机构账号（中国科学院大学+学号）、短信验证码
- 登录成功后脚本会**自动检测到并抓取 Cookie（含 HttpOnly 凭据）**，保存到 `~/.ucas_schedule/`，用户不需要做任何额外操作
- 弹出窗口后脚本会等待（默认 300 秒，可 `--timeout 600` 加长）

密码错误几次会弹文字点选验证码，引导用户按提示点击汉字即可；账号若命中双因子认证按页面提示操作。登录卡住时可建议用户换方式 A。

**切换认证方式**：用户随时可以换（"我想用另一种方式登录"）。方式 B → A：用户自己浏览器登录后复制 Cookie，`save-cookie` 覆盖即可；A → B：直接跑 `login-browser`（`--relogin` 强制重登）。每次换完都重新 `check` 验证。

### 第 3 步：爬取导出

```bash
python3 <CRAWLER> crawl --out <输出目录>
```

脚本行为（无需模型重复实现，但要知道以便向用户解释）：
- 逐学期、逐周（week=1..maxWeek）扫描 `getMyLessons`，按 lessonConfigUuid 去重合并——**不带 week 参数只会返回当前周**，这是该站点的特性
- 每个有课的学期生成 `WakeUp课表_<学年>第<学期>学期.xlsx` 和同名 `.csv`（7 列：课程名称/星期/开始节数/结束节数/老师/地点/周数；周数为 `1-16`、`7-11单`、`2-5、7-20` 这类 WakeUp 规范格式；老师地点缺失填"无"）
- 另存完整原始数据 `schedule_all.json`
- 全学期无课的学期不生成文件（正常现象，如已结课的旧学期数据被清空）

完成后向用户列出生成的文件（带绝对路径链接），并说明 WakeUp 导入方法：App 内 课表 → 导入 → Excel 导入，选 `.csv`（`.xlsx` 亦可），每个学期导一次，在 App 内切换课表。

### 辅助命令

```bash
python3 <CRAWLER> check                      # 单独验证登录态
python3 <CRAWLER> export --json schedule_all.json --out .   # 用原始 JSON 重新生成表格
```

## 常见问题排查

- **`用户信息异常，请登录后再试` / 302**：登录态缺失或失效。Cookie 方式→确认从 Network 请求头整串复制（HttpOnly！）；或换另一种方式重新认证。
- **某学期导出 0 门课 / 无文件**：先确认不是认证问题（`check`），再向用户解释：第 1 周为空、整学期无课程都属正常；真实有课而爬不到才需要排查。
- **登录页密码正确却不动**：多半弹了文字点选验证码（页面内嵌，需人工点击）或双因子认证，让用户在浏览器里完成。
- 重新生成表格不需要重新爬取：用 `export --json` 基于已保存的 `schedule_all.json`。
