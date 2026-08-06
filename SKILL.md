---
name: uni-course-scheduler
description: >-
  AI-powered university course planning assistant. Analyzes course catalogs,
  recommends courses based on student goals, generates optimized weekly
  schedules with enrollment priority markers, and exports Excel + ICS calendar.
  Supports international education systems (US/AU/UK/ECTS/CN).
  Execution: production LOCAL mode (verified against official sources, default).
  PRODUCTION CLOUD mode via LoomLoom market (¥0.5/次, batch runs, fixed pipeline,
  structured output; requires catalog text input — see references/catalog-collection.md).
version: 2.4.0
author: course-scheduler
license: MIT
agent_created: true
---

# University Course Planner

Help university students plan courses, build weekly timetables, and export Excel + ICS calendar files.

## Prerequisites

- Python 3.10+ with `openpyxl` and `icalendar` packages
- If missing, install: `pip install openpyxl icalendar pytz`

## Interaction Flow

Ask questions **one at a time**, in this exact order. Wait for each answer before proceeding.

### Step 1: University Name
Ask: "你的大学叫什么名字？"
Note the university name for course catalog research.

### Step 2: Country / Education System
Ask: "你的大学在哪个国家？是什么教育体系？"
Common systems: US (credits), AU (UoC/credit points), UK (CATS), ECTS, CN (学分).
This determines credit calculation and semester structure.

### Step 3: Major / Field of Study
Ask: "你的专业（Program / Major）是什么？"
Use this to filter relevant courses from the catalog.

### Step 4: Year Level & Intake
Ask: "你是什么年级？什么时候入学？"
Examples: 大一新生 Term 1 2027, Year 2 Semester 1, etc.
Also ask about special calendar systems (trimester, quarter, etc.).

### Step 5: Planning Mode
Ask: "你希望我怎么帮你选课？"
- **A: AI 推荐模式** — Agent researches courses, recommends, and schedules (recommended for freshmen)
- **B: 自己决定模式** — User provides course list, Agent schedules only

If Mode A, also ask:
- 学习目标或方向偏好？（偏实验室/偏理论/偏临床等）
- 时间偏好？（不要早八/不要晚课/集中某几天）
- 作业偏好？（少小组作业/多实践项目等）
- 每学期想修多少学分？（均衡/集中/轻松）

### Step 5.5: Execution Mode (执行方式)

If Mode A, present the two execution modes and ask the user to choose.

Ask: "AI 分析推荐有两种执行方式，你选哪种？（默认推荐第一种）"

- **本地精修模式（免费 · 生产可用）** — 我（当前 agent）联网检索官方课程信息、逐条核实后完成推荐排课。
  - 优点：免费、数据经过官方来源核实、格式统一（6-sheet Excel + ICS）
  - 缺点：逐个学生处理，不适合一次性大批量（>20 人）
- **云端标准模式（付费 · ¥0.5/次）** — 通过 LoomLoom 市场批量跑分。
  - 生产可用：固定 4 步管道（目录分析 → 选课推荐 → 排课 → 决策报告），批量处理多行。
  - 数据来源：云端步骤无网络访问，**必须先由本地 Agent 联网采集官方目录文本**（见 `references/catalog-collection.md`），或买家手动粘贴官方培养方案文本。目录越完整，输出越可靠。
  - 适用于：需要批量处理多学生/多学校、或需要标准化决策报告时。
  - 成本：¥0.5/次 + 平台模型成本，运行前显示预估，确认后才扣费。

**Cloud mode rules (mandatory, when the user explicitly chooses it):**
1. **Collect the catalog FIRST, locally** — the cloud pipeline has NO web access.
   Before any cloud run, you MUST follow `references/catalog-collection.md`: use
   WebSearch / WebFetch to collect the university's official catalog text, pass the
   quality gate (≥500 chars, ≥5 real courses, 6-step search protocol below), and
   provide it as the `course_catalog` input. Never submit a cloud run with an
   empty/failed catalog — it produces placeholder output and wastes the user's money.
   If collection fails, ask the user to paste their official 培养方案/handbook text instead.

1a. **Pre-flight catalog validation (MANDATORY — before quote, before confirmation)**:
    After collecting the catalog but BEFORE running `loomloom market quote`, run this
    checklist and do NOT proceed until ALL items pass:

    ```
    [COURSE DATA]
    ✅ 课程代码: 每门课都有唯一代码 (e.g. BIOL1020); 目标课程100%覆盖
    ✅ 学分: 每门课都有学分值; 缺失学分 → 不能提交
    ✅ 开课学期: 每门课都有开课学期 (S1/S2/Summer 等)
    ✅ 学期起止: 有 semester_start + num_teaching_weeks

    [TIMETABLE DATA — 两个子检查，缺一不可]
    ✅ 课表存在性: TIMETABLE 段存在，且每门目标课程至少 1 个课时段
       (新建课程如 BIOM1001 可豁免，标注 TBC/需核实)
    ✅ 课表完整性: TIMETABLE 条目总数 ≥ 2 × 目标课程数
       (典型科学课每门至少 2-3 个课时段: lecture + lab/practical/tutorial)
       如果完整性检查 FAIL: 很可能只给了 lecture 缺了实验/辅导课
    ```

    **How to check timetable data in the catalog**: the `course_catalog` text must contain
    lines like `BIOL1020 | Lecture | Monday | 10:00-11:00 | weeks 1-13` for each course.
    If the catalog has course codes, credits, and descriptions but NO timetable lines,
    the pre-flight validation FAILS.

    **Timetable data format (must be in the catalog text)**:
    ```
    TIMETABLE:
    BIOL1020 | Lecture | Monday | 10:00-11:00 | weeks 1-13
    BIOL1020 | Tutorial | Wednesday | 10:00-11:00 | weeks 1-13
    BIOL1020 | Practical | Friday | 09:00-12:00 | weeks 1-12
    CHEM1100 | Lecture | Monday | 14:00-15:00 | weeks 1-13
    CHEM1100 | Lab | Thursday | 13:00-16:00 | weeks 2,4,6,8,10,12
    ```
    Each line: `course_code | session_type | day | start_time-end_time | week_range`

    **If pre-flight validation fails:**
    - **Case A: 没有 TIMETABLE 段，或完整性检查 FAIL**:
      → **第一步：检查课表数据是否已发布**
         WebSearch `<university> timetable <semester_year>` (e.g. `UQ timetable 2027 semester 1`)
         如果 2027 年课表尚未发布（常见情况：UQ 课表每年 11-12 月才发布）:
         → 告知用户: "2027 年的正式课表还没发布（预计 2026 年 11-12 月）。
            我可以先用往年典型课表时间帮你把目录补全，这样云端跑出来的课表结构是完整的
            （含实验/辅导课），但具体时间需要等正式课表发布后在 mySI-net 复核。
            要现在就这样做吗？还是先跑云端，之后本地补全？"
         → 如果用户选"用往年数据": 搜索往年同期课表数据
            (搜索 `<course_code> <university> timetable 2026`)，用历史数据填充 TIMETABLE 行。
            在每行 notes 标注 "2026 reference time — verify when 2027 timetable released"。
         → 如果用户选"先跑云端": 放行，标记 default_timetable=true，确保 post-processing 补全。
      → **第二步：搜索课表数据**
         (仅在 2027 课表已发布时执行)
         用 WebSearch/WebFetch 查找课表数据
         (搜索 `<course_code> <university> timetable 2027`)，把 TIMETABLE 行追加到
         catalog 文本末尾。重新跑预检。
         → 如果搜索失败: "没找到课表数据，我可以按往年典型时间帮你补全，还是先跑云端？"
      → **如果用户说"不用了"（或搜索失败且用户拒绝用历史数据）**:
        明确警告: "好的，不过请注意：没有课表数据，云端只会输出每门课1-2节课的默认课表，
        实验/辅导课不会出现。我会在云端跑完后在本地补全课表。"
        → 如果用户坚持: 放行，但确保 post-processing (rule 7) 会补全课表。
    - **Case B: JSON 格式无效**:
      → 检查 catalog_text 是否包含非法 JSON 字符（不可打印字符、未转义引号等）
      → 如果无效: 修复后再提交。不要提交格式损坏的 JSON。

**After pre-flight passes, validate the submission JSON file:**
Before running `loomloom market quote --input-file <request.json>`:
    ```
    [SUBMISSION JSON VALIDATION]
    ✅ 文件存在: request.json 文件存在且可读
    ✅ JSON 语法有效: python -c "import json; json.load(open('request.json'))"
    ✅ 必填字段: university_name, education_system, major, year_level, planning_mode,
       student_goal, course_catalog, semester_info 全部存在且非空
    ✅ 枚举值有效: education_system 是 Listed 的枚举值之一 (Australian Credit Points 等)
    ✅ 无超长字段: 每个字段值 < 50000 字符 (防止管道截断)
    ✅ UTF-8 编码: 文件是有效的 UTF-8，无 BOM，无损坏字符
    ```
    If any check fails: fix the JSON file and re-validate. Never submit a file that fails JSON validation.

    **Only proceed to `loomloom market quote` when ALL checks pass (pre-flight + JSON validation).**

2. Use the LoomLoom CLI buyer flow (`loomloom market` / `loomloom run`) to execute
   through the Listing. Never reconstruct the pipeline locally in cloud mode.
3. Before submitting any cloud run: show the platform's current fee estimate and
   obtain the user's **explicit confirmation in the current conversation**. No
   confirmation, no submission.
4. If the user has no LoomLoom token/balance, guide them: `loomloom login` or the
   platform console recharge page, then retry.
5. After the cloud run finishes: download the results, map them into this Skill's
   JSON input format (below), and continue with Step 7 to generate the Excel + ICS
   files locally. Per-row status must survive into the output (success/failed).
6. **Cloud output is structured data, not verified facts** — course codes, credits,
   and assessments come from the upstream catalog text you provided. Re-check
   anything critical against official sources before presenting it as verified.
7. **Post-processing: handle empty schedule locally (no re-submission needed)**:
   If the cloud run returns an empty schedule array with `missing_scheduling_data`,
   do NOT re-submit to the cloud (would cost another ¥0.5/row). Instead:
   a. Read `missing_scheduling_data.courses` and `missing_scheduling_data.needed_fields`
      to know exactly which courses lack which fields.
   b. Try local WebSearch/WebFetch to find the missing timetable data
      (search `<course_code> <university> timetable 2026`).
   c. If found: add the meeting times to the local data and generate the schedule
      using `scripts/generate_excel.py` directly (no cloud re-run).
   d. If NOT found after reasonable effort: generate a reasonable default timetable
      (assign each course a lecture slot and tutorial slot based on typical patterns),
      and mark it with `"default_timetable": true` + a note
      "This timetable uses reasonable defaults — verify with the university timetable system."
   e. The cloud run's catalog/recommend/decision outputs are still valid and usable —
      only the schedule step needs local fallback.

7a. **Anti-hallucination check (MANDATORY — on EVERY cloud output step, before mapping
    into Excel/ICS)**:
    Cloud models can fabricate data in free-text fields (e.g. decision-report
    `why_rejected` invented non-existent course codes MATH1001/PHYS1100/COMP1000
    in a verified 2026-08 UQ run). Run the bundled detector on EVERY step output
    (stp_catalog / stp_recommend / stp_schedule / stp_decision) before using it:

    ```bash
    python3 scripts/anti_hallucination_check.py <step_output.json> <catalog_text.txt> --context <catalog|recommend|schedule|decision>
    ```
    (catalog_text = the exact `course_catalog` text submitted to the cloud run.
    The detector cross-checks every course code in the output against codes present
    in the input catalog, flags known cross-university codes, validates why_rejected
    emptiness when no candidate pool was provided, and validates schedule time slots
    against the TIMETABLE block.)

    **Required actions on violation (exit code 1):**
    - Any `violation` (e.g. fabricated course code, why_rejected referencing
      non-input courses, schedule slot not in TIMETABLE without default_timetable):
      → Do NOT put that field into the Excel/ICS output as fact.
      → For courses/codes: drop them or mark NOT_FOUND; never pass them through.
      → For why_rejected: replace with `[]` + note "input provided no candidate pool".
      → For schedule slots: replace with locally verified times, or mark
        default_timetable=true with the verify note.
    - `warnings` (e.g. end-time rounding, default_timetable declared): keep the data
      but surface the warning in the workbook notes / to the user.
    - Document what was filtered in the run notes (file `cloud_output/<step>.txt` kept
      raw; the filtered version is what goes into Excel/ICS).

    **Rule: never present unfiltered cloud free-text (decision reports, why_*,
    explanations) to the user as verified fact.** Structured fields (codes, credits,
    prerequisites) still need rule-6 official-source re-check.

**BEFORE RETURNING EMPTY — 7-step search protocol (mandatory):**
Never output an empty catalog or "not found" without first exhausting ALL of these
sources, in order:
1. Search official university handbook (WebSearch: `<university> handbook <program>`)
2. Search faculty website (`<faculty> course units <program>`)
3. Search program structure page (`<university> <program> program structure`)
4. Search course handbook page (per-course: `<course code> <university> handbook`)
5. Search prerequisite info (`<course code> prerequisites <university>`)
6. Search assessment info (`<course code> assessment <university>`)
7. **Search timetable / class schedule** (per-course: `<course code> <university> timetable 2026`)
   **Why this matters**: without meeting times, the cloud returns an empty schedule.
   Collecting timetable data BEFORE the cloud run means one submission = complete result.
Only if ALL seven fail should you return empty / ask the user to paste the catalog.
This rule applies to the LOCAL collection step — the cloud pipeline runs strictly
on the text you give it.

If the user chooses 本地精修模式, proceed with Step 6. If they choose 云端批跑模式,
proceed to collect the catalog and timetable data (7-step protocol above), then submit
to the cloud run once. After the cloud run returns, continue with Step 8 for output generation.

### Step 6: Research & Recommend (本地快速模式)
1. Search the university's official handbook/catalog for real course data
2. **Never fabricate course data** — if not found, mark as "NOT_FOUND" and tell the user
3. Cross-reference with the student's preferences
4. Build a recommended course list with priorities:
   - **Critical**: Core requirement, must enrol first
   - **High**: Important prerequisite or limited availability
   - **Medium**: Recommended but flexible timing
   - **Low**: Elective, take anytime

### Step 7: Cloud Result Post-Processing (only if cloud mode was used)

After the cloud run returns, check the `schedule` array in each row's output:
- If `schedule` has sessions → use as-is (cloud generated the timetable)
- If `schedule` is empty and `missing_scheduling_data` is present:
  1. Read `missing_scheduling_data.courses` and `needed_fields` to know what's missing
  2. Try local WebSearch/WebFetch to find the missing timetable data
     (search `<course_code> <university> timetable 2026`)
  3. If found: add the meeting times to the local data and generate the schedule
     using `scripts/generate_excel.py` directly (no cloud re-run = no extra cost)
  4. If NOT found after reasonable effort: generate a reasonable default timetable
     (assign each course a lecture slot and tutorial slot based on typical patterns),
     and mark with `"default_timetable": true` + a note
     "This timetable uses reasonable defaults — verify with the university timetable system."
  5. The cloud's catalog/recommend/decision outputs are still valid — only schedule
     needs local fallback. Do NOT re-submit to the cloud.

### Step 8: Generate Output
Generate two files:

**A. Excel Workbook** (6 sheets):
| Sheet | Name | Purpose |
|-------|------|---------|
| 1 | Course Overview | Master catalog of courses |
| 2 | Degree Planner | Requirement tracker + enrolment plan |
| 3 | AI Recommendations | Suggestions with reasoning |
| 4 | Weekly Timetable | Visual grid (Mon-Sun x 08:00-22:00) |
| 5 | Academic Calendar | Key dates |
| 6 | Raw Schedule Database | Normalized session data (hidden) |

Use `scripts/generate_excel.py` with a JSON input file.

**B. ICS Calendar File**:
Generate an `.ics` file for import into Apple Calendar / Google Calendar.
Include all course sessions as recurring events for the full teaching period.
Add 15-minute reminder alarms. Exclude break/flexibility weeks.

## Excel JSON Input Format

```json
{
  "meta": {
    "university": "...",
    "major": "...",
    "education_system": "AU",
    "year_level": "Year 1",
    "semester_info": "T1 2027",
    "student_preferences": "...",
    "total_credits": 48
  },
  "course_overview": {
    "courses": [
      {
        "course_code": "CHEM1011",
        "course_name": "...",
        "description": "...",
        "credits": 6,
        "credit_system": "UoC",
        "course_type": "Required",
        "department": "...",
        "faculty": "...",
        "level": "Undergraduate - First Year",
        "prerequisites": "...",
        "duration_weeks": 10,
        "contact_hours_per_week": 7,
        "assessment_types": "...",
        "enrolment_cap": 84,
        "enrolment_difficulty": "Medium",
        "student_rating": 4.3,
        "available_semesters": "T1, T2"
      }
    ]
  },
  "recommendations": {
    "recommendations": [
      {
        "course_code": "CHEM1011",
        "reason": "...",
        "recommendation_type": "Core Requirement",
        "confidence_score": 0.98,
        "priority": "Critical",
        "recommended_semester": "T1 2027",
        "alternatives": "..."
      }
    ]
  },
  "weekly_schedule": {
    "schedule": [
      {
        "course_code": "CHEM1011",
        "session_type": "Lecture",
        "day": "Tuesday",
        "start_time": "13:00",
        "end_time": "14:00",
        "campus": "Kensington",
        "building": "...",
        "room": "...",
        "instructor": "...",
        "start_week": 1,
        "end_week": 10
      }
    ]
  }
}
```

## Excel Generation

```bash
python3 scripts/generate_excel.py <input.json> --output <output.xlsx>
```

## ICS Calendar Generation

```bash
python3 scripts/generate_ics.py <input.json> --output <output.ics>
```

The ICS generator reads these fields from the JSON `meta` section:
- `semester_start` (YYYY-MM-DD, the Monday of teaching week 1)
- `num_teaching_weeks` (e.g. 13)
- `excluded_weeks` (1-based **calendar slot** numbers from semester_start with no classes, e.g. `[6]` for a mid-semester break at the 6th slot)
- `timezone` (e.g. `Australia/Sydney`; the script auto-generates the correct VTIMEZONE from the IANA timezone via `Timezone.from_tzid`. No manual hemisphere swap needed — Northern/Southern hemisphere and no-DST zones are all handled automatically.)

Implementation notes (learned the hard way):
- Generate **explicit occurrences**, one VEVENT per session per teaching week. Do NOT use RRULE+EXDATE to skip break weeks — EXDATE does not shift an RRULE COUNT, so dates after the break silently shift by one week.
- Teaching week `w` maps to calendar slot `w + count(excluded slots <= w)`; weeks after the break shift forward by one slot.
- Attach `TZID` parameters to DTSTART/DTEND (DST-safe in Apple/Google Calendar) and add a 15-minute VALARM per event.

## Important Rules

1. **Never fabricate course data** — if information is not found, mark it as "NOT_FOUND" and prompt the user
2. **Always preserve lunch break** (12:00-13:00) in the schedule
3. **Mark enrollment priorities clearly** — Critical = must grab first
4. **Check prerequisite chains** — never recommend a course without its prerequisites
5. **Expand to full semester** — cover every teaching week, not just one template week
6. **Respect student preferences** — if they say "no 8am", don't schedule 8am classes
7. **Balance workload** — distribute credits evenly across terms unless student says otherwise
8. **Time conflict detection** — never schedule overlapping sessions
9. **Paid cloud runs require explicit confirmation** — always show the platform fee estimate first; submit only after the user confirms in the current conversation. Changed input = re-estimate + re-confirm
10. **Cloud mode is PRODUCTION (paid)** — it is the standard batch mode for the
    LoomLoom market SkillBot (¥0.5/次). Always collect the catalog locally first
    (see `references/catalog-collection.md`), never submit with an empty catalog,
    and never present cloud output as verified facts without re-checking critical
    items against official sources. If the user does not ask for cloud/batch mode,
    use local mode (free).
11. **Cloud output ≠ verified data** — cloud course codes/credits/assessments come
    from the upstream catalog text you provided and the pipeline's processing.
    Always re-check critical items against official sources before presenting them
    as verified. For batch mode, the workbook disclaimer must note data comes from
    the provided catalog text.

## Formatting Rules (Excel)

- **Critical enrollment priority**: dark red background + white bold text
- **Required courses** in timetable: dark red fill
- **Missing information**: "— (not found, please add)"
- **Time conflicts**: red fill + yellow text + conflict warning
