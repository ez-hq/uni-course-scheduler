---
name: uni-course-scheduler
description: >-
  AI-powered university course planning assistant. Analyzes course catalogs,
  recommends courses based on student goals, generates optimized weekly
  schedules with enrollment priority markers, and exports a 6-sheet Excel
  workbook. Supports international education systems (US/AU/UK/ECTS/CN).
version: 1.0.0
author: course-scheduler
license: MIT
---

# University Course Planner Skill

## What This Skill Does

Helps university students (especially freshmen) plan their course selection and weekly timetable. It walks the student through a guided conversation, analyzes their university's course catalog, recommends courses based on their goals, and generates a comprehensive Excel workbook with:

- Full course overview (6 sheets, 100+ fields)
- Optimized weekly timetable with conflict detection
- AI recommendations with reasoning and alternatives
- Enrollment priority markers (Critical/High/Medium/Low)
- Degree requirement tracking and GPA calculation
- Academic calendar with .ics export support

## Prerequisites

### Required
- Python 3.10+ (for Excel generation script)
- `openpyxl` Python package (install: `pip install openpyxl`)

### Optional (for cloud-powered batch processing)
- LoomLoom CLI installed (`loomloom` command available)
- LoomLoom API token configured via `loomloom login`

### If LoomLoom is NOT installed
The skill still works in **local-only mode**: the Agent uses its own reasoning to analyze courses and generate schedules, then calls the Excel generator script locally. Cloud batch processing is skipped.

## Quick Start (for Agents reading this)

1. **Read** `references/interaction-flow.md` for the 7-step guided conversation script.
2. **Follow** the interaction flow exactly — ask questions one at a time, in order.
3. **After** collecting all inputs, decide:
   - If LoomLoom CLI is available → use cloud execution (see `references/loomloom-setup.md`)
   - If not → use local reasoning mode (Agent does the analysis itself)
4. **Generate** the Excel file using `scripts/generate_excel.py`.
5. **Validate** the output using `scripts/validate_schedule.py`.
6. **Deliver** the .xlsx file to the user.

## Interaction Flow Summary

| Step | What to Ask | Why |
|------|------------|-----|
| 1 | University Name | To search for course catalog and credit system |
| 2 | Country / Education System | Different countries use different credit frameworks |
| 3 | Major / Field of Study | To filter relevant courses |
| 4 | Year Level | To match course difficulty and prerequisites |
| 5 | Planning Mode (AI recommend or user decided) | Determines whether AI picks courses or just schedules |
| 5a | If AI: Goals + Schedule preferences | Drives recommendation logic |
| 6 | Semester Info | Needed to expand schedule to full semester dates |
| 7 | Generate Excel | Final output delivery |

See `references/interaction-flow.md` for the full script with example prompts.

## Excel Output

The generated workbook has 6 sheets matching the schema in `references/excel-output-spec.md`:

| Sheet | Name | Purpose |
|-------|------|---------|
| 1 | Course Overview | Master catalog of all available courses |
| 2 | Degree Planner | Requirement tracker + semester enrolment plan |
| 3 | AI Recommendations | AI-generated suggestions with reasoning |
| 4 | Weekly Timetable | Visual grid (Mon-Sun x 08:00-22:00) |
| 5 | Academic Calendar | Key dates and events |
| 6 | Raw Schedule Database | Normalized session data (hidden by default) |

### Key Formatting Rules
- **Critical enrollment priority** courses: dark red background + white bold text
- **Required courses** in timetable: dark red fill
- **Missing information**: displayed as "— (not found, please add)"
- **Time conflicts**: red fill + yellow text + conflict warning

## Cloud Execution (LoomLoom)

If LoomLoom CLI is available, the skill uses a 3-step cloud pipeline:

```
Step 1: Course Catalog Analysis (stp_catalog)
    → Parses uploaded catalog, classifies courses, extracts metadata
Step 2: Course Recommendation (stp_recommend)
    → Generates personalized recommendations based on goals
Step 3: Schedule Generation (stp_schedule)
    → Optimizes timetable, detects conflicts, assigns priorities
```

The cloud output JSON is then passed to `scripts/generate_excel.py` to produce the Excel file.

See `references/cloud-output-format.md` for the expected JSON structure.

## Local Validation

After generating the Excel file, run:

```bash
python3 scripts/validate_schedule.py <excel_file.xlsx>
```

This checks:
- Time conflict detection (no overlapping sessions)
- Credit total validation
- Required course coverage
- Schedule density and lunch break preservation
- Missing field detection

See `references/local-validation.md` for the full validation contract.

## File Structure

```
uni-course-scheduler/
├── SKILL.md                          (this file)
├── agents/
│   └── openai.yaml                   (OpenAI agent config)
├── references/
│   ├── interaction-flow.md           (7-step guided conversation)
│   ├── excel-output-spec.md          (6-sheet Excel schema)
│   ├── cloud-output-format.md        (expected JSON from cloud)
│   ├── local-validation.md           (validation rules)
│   └── loomloom-setup.md             (LoomLoom CLI setup guide)
└── scripts/
    ├── generate_excel.py             (JSON → Excel workbook)
    ├── validate_schedule.py          (Excel → validation report)
    └── local_audit.py                (audit cloud output before Excel)
```

## Supported Platforms

- WorkBuddy (primary)
- Claude Code
- Codex
- OpenClaw

## Important Rules

1. **Never fabricate course data** — if information is not found, mark it as "NOT_FOUND" and prompt the user to provide it.
2. **Always preserve lunch break** (12:00-13:00) in the schedule.
3. **Mark enrollment priorities clearly** — Critical = must grab first, marked in red.
4. **Check prerequisite chains** — never recommend a course without its prerequisites.
5. **Expand to full semester** — the Excel must cover every teaching week, not just one week template.
6. **Respect the student's preferences** — if they say "no 8am", don't schedule 8am classes.
