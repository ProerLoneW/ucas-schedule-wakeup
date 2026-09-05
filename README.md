# UCAS Schedule WakeUp Exporter

国科大（中国科学院大学）课表导出工具：通过 [kb.mooc.ucas.edu.cn](https://kb.mooc.ucas.edu.cn/res/pc/curriculum/schedule.html)（超星课表国科大部署）的身份认证，爬取**所有学年学期**的课程，按学期分别导出 [WakeUp 课程表](https://wakeup.fun) App 可直接导入的 Excel/CSV 文件。

提供两种使用方式：**作为 Agent Skill 使用**（对话式引导，推荐）或**直接运行 Python 脚本**。

## 工作原理

- 课表数据接口 `GET /pc/curriculum/getMyLessons` 需要登录态（Cookie 中含 HttpOnly 凭据）
- 该接口**不带 `week` 参数时只返回当前周**的课程，因此工具会逐周（week=1..maxWeek）扫描再按课程唯一 ID 去重，得到整学期完整课表
- 导出为 WakeUp 官方 7 列导入格式（见 [官方文档](https://wakeup.fun/doc/import_from_csv.html)）：`课程名称 | 星期 | 开始节数 | 结束节数 | 老师 | 地点 | 周数`，周数支持 `1-16`、`7-11单`、`2-5、7-20` 等写法，老师/地点缺失自动填"无"

## 仓库结构

```
ucas-schedule-wakeup/
├── README.md                        # 本文件
├── login_browser.py                 # 入口：弹浏览器登录学习通，保存登录态
├── crawl.py                         # 入口：爬取全部学期并导出 WakeUp 文件
├── requirements.txt                 # Python 依赖
├── .env.example                     # 环境变量模板（复制为 .env 填入 Cookie）
└── ucas-schedule-export/            # Agent Skill（也可独立使用）
    ├── SKILL.md                     # Skill 引导逻辑（供 Agent 读取）
    └── scripts/
        └── ucas_crawler.py          # 核心工具（子命令：doctor/login-browser/
                                    #   save-cookie/check/crawl/export）
```

登录态文件统一保存在 `~/.ucas_schedule/`，两种使用方式共享，跨会话可复用。

## 方式一：作为 Agent Skill 使用（推荐）

适用于 [ZCode](https://zcode.zhipuai.com) 等支持 Agent Skill 机制的 AI 编程工具。

### 安装 Skill

```bash
# 用户级安装（所有项目可用）
cp -r ucas-schedule-export ~/.agents/skills/

# 或项目级安装（仅当前仓库可用）
cp -r ucas-schedule-export <你的项目>/.zcode/skills/
```

安装 Python 依赖：

```bash
pip3 install openpyxl playwright
python3 -m playwright install chromium   # 仅浏览器登录方式需要
```

### 与 Agent 交互

安装后**新开一个对话**，正常描述需求即可触发，例如：

- "帮我爬一下我国科大的课表，导成 WakeUp 能导入的格式"
- "我的 kb.mooc.ucas.edu.cn 课表想导出 Excel"

也可以用斜杠命令强制触发：`/ucas-schedule-export 导出本学期课表`

Skill 会引导你完成认证（**二选一，随时可切换**）：

| 认证方式 | 过程 | 适合场景 |
|---|---|---|
| **Cookie 导入** | Agent 教你在浏览器 F12 → Network 里复制整串 Cookie，粘贴给 Agent 验证 | 不想装 playwright；已在浏览器登录过 |
| **学习通账号登录** | Agent 运行脚本弹出浏览器，你扫码或输入账号密码登录，脚本自动检测并抓取 Cookie | 想要一次登录长期复用 |

认证通过后 Agent 自动爬取全部学期并导出文件，最后把文件路径列表给你。

## 方式二：直接运行 Python 脚本

### 环境准备

```bash
pip3 install -r requirements.txt
python3 -m playwright install chromium   # 仅浏览器登录方式需要
```

### 认证方式 A：Cookie + 环境变量（`.env`）

```bash
cp .env.example .env        # 然后编辑 .env，填入 UCAS_COOKIE
python3 crawl.py            # 或 export UCAS_COOKIE="..." 后直接运行
```

**Cookie 获取步骤**：

1. 浏览器打开 [课表页](https://kb.mooc.ucas.edu.cn/res/pc/curriculum/schedule.html)，按提示登录（跳转到 passport 统一认证，扫码/手机号+学习通密码/机构账号"中国科学院大学"+学号均可）
2. 登录跳回课表页后，按 F12 → **Network** 标签 → 刷新页面
3. 点击任意一个发往 `kb.mooc.ucas.edu.cn` 的请求（如 `getMyLessons`）
4. 在 **Request Headers** 里找到 `Cookie:` 一行，复制**整串值**填入 `.env`

> ⚠️ 必须从 Network 请求头复制。控制台 `document.cookie` 读不到 HttpOnly 凭据（如 vc3），缺了会被服务器 302 踢回登录页。

### 认证方式 B：弹浏览器登录学习通

```bash
python3 login_browser.py           # 弹出浏览器，扫码/输密码登录
python3 login_browser.py --relogin # 登录态失效时强制重登
```

登录成功自动保存到 `~/.ucas_schedule/`，之后直接运行 `crawl.py` 即可，无需重复登录。

### 爬取导出

```bash
python3 crawl.py                  # 登录态自动取自 .env / 环境变量 / 已保存文件
python3 crawl.py --out ./output   # 指定输出目录
python3 crawl.py --no-merge       # 教务课程不合并，更细粒度
```

输出（每个**有课**的学期一组文件）：

```
WakeUp课表_2026-2027第1学期.xlsx   # Excel 版，方便自行查看修改
WakeUp课表_2026-2027第1学期.csv    # WakeUp 导入用（App 内 课表→导入→Excel导入）
schedule_all.json                 # 接口完整原始数据留档
```

**导入 WakeUp**：App 内 课表 → 导入 → Excel 导入，选择 `.csv` 文件（`.xlsx` 亦可）；每个学期导一次，在 App 内切换课表。

## 常见问题

- **提示"用户信息异常，请登录后再试" / 登录态无效**：Cookie 复制不完整（缺 HttpOnly 项）或已过期。重新从 Network 请求头整串复制，或改用 `login_browser.py`。
- **某个学期没有生成文件**：该学期全学期无课程（接口逐周扫描均为 0）属正常；第 1 周课少/为空也是正常现象。
- **登录页密码正确但没反应**：页面内嵌文字点选验证码或双因子认证，在浏览器里按提示完成即可。
- **出现 SSL 证书警告并自动降级**：macOS 自带 Python 缺根证书所致，可运行 `/Applications/Python*/Install\ Certificates.command` 修复；不影响功能。
- **只想重新生成表格不想重新爬**：`python3 ucas-schedule-export/scripts/ucas_crawler.py export --json schedule_all.json`

## 补充说明

