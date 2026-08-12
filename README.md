# uni-course-scheduler 大学选课排课助手

> AI-powered university course planning assistant — analyze course catalogs, get smart course recommendations, build conflict-free weekly timetables, and export Excel + ICS calendar files.
>
> AI 大学选课规划助理：分析官方课程目录，智能推荐课程，生成无冲突的每周课表，并导出 Excel + ICS 日历文件。

[![License](https://img.shields.io/badge/license-Personal%20Use-blue.svg)](LICENSE)

---

## What it does 功能

`uni-course-scheduler` is an AI Skill that helps university students plan their semester:

- **Catalog analysis 目录分析** — parses a school's official course catalog (program handbook text) into structured course data: codes, names, credits/points, prerequisites, term availability, and timetable slots.
- **Course recommendation 课程推荐** — recommends courses matched to the student's goals (e.g. GPA-first, ML specialization), flags enrollment priority, and validates prerequisites.
- **Weekly schedule 每周排课** — builds a conflict-free weekly timetable from the real timetable, respecting preferences like "no 8am classes" or "keep 12:00–13:00 free".
- **Excel export 表格导出** — generates a 6-sheet (local) or 7-sheet (cloud) Excel workbook, including `Weekly Timetable`, `Degree Planner`, and (cloud) `Degree Audit & Risk Assessment`.
- **ICS calendar 日历导出** — exports a downloadable `.ics` file for Apple Calendar / Google Calendar (cloud mode).

Supports **international credit systems** (US / AU / UK / ECTS / CN / HK / SG) and **undergraduate and graduate** students.

支持国际学分体系（美/澳/英/ECTS/中/港/新），覆盖本科与研究生。

---

## How it works (two modes) 两种模式

This Skill ships in two execution modes:

### 1. Local personal mode (free) 本地个人版（免费）
The agent itself researches the official course catalog, verifies it, and produces a 6-sheet Excel + a textual weekly overview.
- Free, single-school, single-student, personal use only.
- No batch pipeline, no standardized decision report, no downloadable `.ics` calendar, no deep degree audit.

本地个人轻量版：免费，仅限单人单校，无批量管道、无标准决策报告、无可下载 `.ics` 日历、无深度毕业审计。

### 2. Cloud standard mode (paid) 云端标准模式（付费）
A fixed production-grade cloud pipeline (`catalog → recommend → schedule → decision → ICS → degree-audit`) processes the catalog text and returns:
- Batch / multi-school / multi-student execution.
- Standardized 8-dimension **Decision Report** 决策报告.
- **Degree Audit & Risk Assessment** sheet (prerequisite chains, credit progress, overload risks).
- **Downloadable `.ics` calendar** (university term dates, weekly recurrence, correct IANA timezone).

云端标准模式：固定生产级云端管道（目录→推荐→排课→决策→ICS→毕业审计），支持批量/多校/多人提交，含标准决策报告、毕业审计、可下载 `.ics` 日历。

> **Offline-cloud note 云端离线说明:** Like most hosted LLM pipelines, the cloud executor has no internet access. The local agent (or the user) must first paste the official catalog text. The pipeline never fabricates data — anything not found is reported honestly as missing.
>
> 云端执行器无联网能力，需先提供学校官方目录文本；管道绝不编造数据，输入中缺失的信息会如实标注为缺失。

---

## Cloud platform (choose your payment) 云端平台（按支付方式选择）

The cloud command is platform-agnostic; the platform is chosen by your payment method:

| Your payment 支付方式 | Platform 平台 | Price 价格 | Server | Token env |
|---|---|---|---|---|
| 中国大陆支付 Alipay/WeChat/UnionPay | 胜算云 Shengsuanyun | ¥0.5/次 | `loomloom.shengsuanyun.com` | `LOOMLOOM_TOKEN_SHENGSUANYUN` |
| 国外信用卡 Intl. credit card (Visa/MC) | CogFoundry | $0.10/次 | `loomloom.cogfoundry.ai` | `LOOMLOOM_TOKEN_COGFOUNDRY` |

On first cloud use, the assistant asks which payment you use and configures the matching server. See `references/loomloom-setup.md` for full setup.

首次云端使用时，助手会询问你的支付方式并配置对应平台，详见 `references/loomloom-setup.md`。

---

## Install 安装

### Prerequisites 前置要求
- An AI assistant that supports custom Skills (e.g. WorkBuddy / Claude Code / Codex / Cursor-style agent). 支持自定义 Skill 的 AI 助手。
- For **cloud mode**: a platform account + balance 云端模式需平台账户与余额.

### Option A — drag & drop (friendliest) 方式A：拖拽安装（最省心）
1. Download the `uni-course-scheduler.zip` release. 下载 ZIP 安装包。
2. Drag the ZIP into your AI chat and say: **"Please install this planning Skill for me."** 拖入 AI 对话并请求安装。
3. The AI agent installs it and you can start. 安装后即可使用。

### Option B — manual install 方式B：手动安装
```bash
# 1. Unzip the package 解压
mkdir -p ~/.workbuddy/skills && cd ~/.workbuddy/skills
unzip /path/to/uni-course-scheduler.zip

# 2. Verify the folder structure 确认结构
ls ~/.workbuddy/skills/uni-course-scheduler   # SKILL.md, scripts/, references/
```

### Cloud setup (optional, for cloud mode) 云端配置（可选）
```bash
# Install the LoomLoom CLI (buyer flow) 安装 LoomLoom CLI
# See references/loomloom-setup.md for full instructions
loomloom login          # sign in with your key 登录
loomloom market list    # verify connectivity 验证连通
```

---

## Usage 使用方法

### Plan a course schedule 规划课程
1. Tell the agent your school, major, year, and goals (e.g. "GPA first, prefer no 8am classes"). 告知学校/专业/年级/目标。
2. Provide the official catalog text — let the local agent fetch it, or paste it. 提供官方目录文本。
3. Choose a mode 选择模式: **Local (free)** or **Cloud (paid)**.
4. Confirm the fee (cloud) and receive 确认费用后获得: Excel workbook (`.xlsx`) + downloadable `.ics` (cloud).

### Collect a course catalog 收集课程目录
See `references/catalog-collection.md` for a reusable prompt and quality gates (≥500 chars, ≥5 courses, official source).

---

## Project structure 项目结构

```
uni-course-scheduler/
├── SKILL.md                      # Main skill instructions (both modes) 主指令
├── README.md                     # This file 本说明
├── LICENSE                       # Personal-use license 个人使用许可
├── agents/
│   └── openai.yaml               # OpenAI / compatible agent metadata
├── references/
│   ├── catalog-collection.md     # How to collect & verify catalogs 目录采集
│   ├── excel-output-spec.md      # Excel workbook schema
│   ├── cloud-output-format.md    # Cloud pipeline output format
│   ├── interaction-flow.md       # End-to-end user flow
│   ├── local-validation.md       # Local output validation rules
│   └── loomloom-setup.md         # Cloud setup (dual-platform) 双平台配置
└── scripts/
    ├── generate_excel.py         # Build the Excel workbook from JSON
    ├── generate_ics.py           # Render .ics from cloud ICS data
    ├── anti_hallucination_check.py  # Anti-hallucination validation 反幻觉校验
    ├── validate_schedule.py      # Schedule conflict validation
    ├── local_audit.py            # Local output audit
    └── check_github_downloads.py # Download stats helper
```

---

## Quality & anti-hallucination 质量与反幻觉

- **No fabrication 不编造:** data not present in the input is labeled `NOT_FOUND` / missing — never invented.
- **Honest empty 诚实空结果:** for garbage/empty input returns an honest empty result with guidance, not a hallucination.
- **Validation 校验:** local scripts run schedule-conflict checks and an anti-hallucination audit on every output.

---

## License 许可

**Personal Use License — No Commercial Redistribution.** Free for personal use; commercial redistribution or re-listing on paid marketplaces requires the author's written permission. See [LICENSE](LICENSE) (Chinese text prevails).

个人使用许可，禁止商业再分发；商业用途或付费市场上架需作者书面授权。

---

## Support / Feedback 支持与反馈

- Author / 作者: ez-hq
- Found a bug or have a catalog request? Open an issue in the repository / 欢迎提交 Issue。