# uni-course-scheduler 大学选课排课助手

> AI 大学选课规划助理：分析官方课程目录，智能推荐课程，生成无冲突的每周课表，并导出 Excel + ICS 日历文件。

[![License](https://img.shields.io/badge/license-Personal%20Use-blue.svg)](LICENSE)

**语言 / Language:** [English](README.md) | [中文](README.zh-CN.md)

---

## 功能

`uni-course-scheduler` 是一个帮助大学生规划学期课程的 AI Skill：

- **目录分析** — 把学校官方课程目录（专业培养方案文本）解析成结构化课程数据：课程代码、名称、学分、先修要求、开课学期、上课时间。
- **课程推荐** — 根据学生目标（如 GPA 优先、机器学习方向）推荐课程，标注选课优先级，校验先修要求。
- **每周排课** — 基于真实课表生成无时间冲突的每周课程表，尊重诸如"不要早八"、"午休 12:00–13:00 保留"等偏好。
- **表格导出** — 生成 6 个 sheet（本地版）或 7 个 sheet（云端版）的 Excel 工作簿，含 `Weekly Timetable`（每周课表）、`Degree Planner`（学位规划）以及云端版 `Degree Audit & Risk Assessment`（毕业审计与风险评估）。
- **日历导出** — 导出可下载的 `.ics` 文件，可直接导入 Apple 日历 / Google 日历（云端模式）。

支持**国际学分体系**（美/澳/英/ECTS/中/港/新），覆盖**本科与研究生**。

---

## 两种模式

本 Skill 提供两种执行模式：

### 1. 本地个人版（免费）
由本地 agent 自行检索并核实官方课程目录，生成 6-sheet Excel + 文字版每周概览。
- 免费，仅限单人单校，个人使用。
- 无批量管道、无标准决策报告、无可下载 `.ics` 日历、无深度毕业审计。

### 2. 云端标准模式（付费）
一条固定、生产级的云端管道（`目录 → 推荐 → 排课 → 决策 → ICS → 毕业审计`）处理目录文本并返回：
- 批量 / 多校 / 多人一次提交执行。
- 标准化 **8 维决策报告**。
- **毕业审计与风险评估** sheet（先修链、学分进度、课程超载风险）。
- **可下载的 `.ics` 日历**（精确的学期日期、每周重复、正确的 IANA 时区）。

> **云端离线说明：** 与大多数托管 LLM 管道一样，云端执行器无联网能力。本地 agent（或用户）必须先粘贴官方目录文本。管道绝不编造数据——输入中缺失的信息会如实标注为缺失。

---

## 云端平台（按支付方式选择）

云端命令与平台无关，实际平台由你的支付方式决定：

| 支付方式 | 平台 | 价格 | Server | Token 环境变量 |
|---|---|---|---|---|
| 中国大陆支付（支付宝/微信/银联） | 胜算云 | ¥0.5/次 | `loomloom.shengsuanyun.com` | `LOOMLOOM_TOKEN_SHENGSUANYUN` |
| 国外信用卡（Visa/Mastercard） | CogFoundry | $0.10/次 | `loomloom.cogfoundry.ai` | `LOOMLOOM_TOKEN_COGFOUNDRY` |

首次使用云端时，助手会询问你的支付方式并配置对应的 server。完整配置见 `references/loomloom-setup.md`。

---

## 安装

### 前置要求
- 支持自定义 Skill 的 AI 助手（如 WorkBuddy / Claude Code / Codex / Cursor 等）。
- **云端模式**需要平台账户与余额。

### 方式 A：拖拽安装（最省心）
1. 下载 `uni-course-scheduler.zip` 安装包。
2. 把 ZIP 拖入 AI 对话并说：**「请帮我安装这个排课 Skill」**。
3. AI agent 安装完成即可使用。

### 方式 B：手动安装
```bash
# 1. 解压安装包
mkdir -p ~/.workbuddy/skills && cd ~/.workbuddy/skills
unzip /path/to/uni-course-scheduler.zip

# 2. 确认目录结构
ls ~/.workbuddy/skills/uni-course-scheduler   # 应包含 SKILL.md, scripts/, references/
```

### 云端配置（可选，云端模式需要）
```bash
# 安装 LoomLoom CLI（买家流程）
# 完整说明见 references/loomloom-setup.md
loomloom login          # 用你的密钥登录
loomloom market list    # 验证连通性
```

---

## 使用方法

### 规划一门课程的课表
1. 告诉 agent 你的学校、专业、年级、目标（如"GPA 优先，不要早八"）。
2. 提供官方目录文本——可让本地 agent 检索，也可自行粘贴。
3. 选择模式：**本地版（免费）** 或 **云端版（付费）**。
4. 确认费用（云端）后获得：Excel 工作簿（`.xlsx`）+ 可下载的 `.ics` 日历（云端）。

### 收集课程目录
见 `references/catalog-collection.md`，含可复用提示词与质量门槛（≥500 字符、≥5 门课程、官方来源）。

---

## 项目结构

```
uni-course-scheduler/
├── SKILL.md                      # 主指令（两种模式）
├── README.md                     # 本说明（英文版）
├── README.zh-CN.md               # 本说明（中文版）
├── LICENSE                       # 个人使用许可
├── agents/
│   └── openai.yaml               # OpenAI / 兼容 agent 元数据
├── references/
│   ├── catalog-collection.md     # 如何采集与核验目录
│   ├── excel-output-spec.md      # Excel 工作簿 schema
│   ├── cloud-output-format.md    # 云端管道输出格式
│   ├── interaction-flow.md       # 端到端用户流程
│   ├── local-validation.md       # 本地输出校验规则
│   └── loomloom-setup.md         # 云端配置（双平台）
└── scripts/
    ├── generate_excel.py         # 从 JSON 生成 Excel 工作簿
    ├── generate_ics.py           # 从云端 ICS 数据渲染 .ics
    ├── anti_hallucination_check.py  # 反幻觉校验
    ├── validate_schedule.py      # 课表冲突校验
    ├── local_audit.py            # 本地输出审计
    └── check_github_downloads.py # 下载统计助手
```

---

## 质量与反幻觉

- **不编造：** 输入中不存在的数据标注为 `NOT_FOUND` / 缺失，绝不虚构。
- **诚实空结果：** 面对垃圾/空输入，返回带指引的诚实空结果，而非道歉或幻觉。
- **校验：** 本地脚本对每次输出运行课表冲突检查与反幻觉审计。

---

## 许可

**个人使用许可 — 禁止商业再分发。** 个人使用免费；商业再分发或在付费市场上架需作者书面授权。见 [LICENSE](LICENSE)（以中文文本为准）。

---

## 支持与反馈

- 作者：ez-hq
- 发现 bug 或有课程目录需求？欢迎提交 Issue。