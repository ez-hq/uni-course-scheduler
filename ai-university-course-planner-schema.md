# AI University Course Planner — Excel Workbook Schema Specification

**Version:** 2.0.0

**Status:** Final

**Target Platform:** Microsoft Excel (.xlsx), Google Sheets, LibreOffice Calc

**Encoding:** UTF-8

**Language:** English (international); field labels localizable via i18n key column

---

## Table of Contents

1. Workbook Metadata
2. Global Design Standards
3. Worksheet 1 — Course Overview
4. Worksheet 2 — Degree Planner
5. Worksheet 3 — AI Recommendations
6. Worksheet 4 — Weekly Timetable
7. Worksheet 5 — Academic Calendar
8. Worksheet 6 — Raw Schedule Database
9. Cross-Worksheet Relationships
10. Enum Definitions
11. Color Rules
12. Formula Rules
13. Timetable Layout Rules
14. Calendar Export Rules
15. Print Settings
16. Data Validation Dictionary
17. Conditional Formatting Dictionary
18. Versioning & Migration Strategy

---

## 1. Workbook Metadata

| Property | Value |
|---|---|
| File name pattern | `{UniversityCode}_{DegreeCode}_{StudentID}_CoursePlanner.xlsx` |
| Default sheets | 6 |
| Hidden sheets | 0 (all visible by default; Raw Schedule Database may be hidden per institution preference) |
| Protected structure | Yes (workbook structure locked; individual sheet protection optional) |
| Password protection | None (enforced via Excel sheet protection without password; institution may add) |
| Shared workbook | Supported (co-authoring via OneDrive / Google Drive) |
| Macro-enabled | No (.xlsx format; no VBA dependencies) |
| External data connections | None (self-contained; AI generation populates on creation) |
| Minimum Excel version | Excel 2016+ / Google Sheets (current) / LibreOffice 8+ |

---

## 2. Global Design Standards

### 2.1 Naming Conventions

- **Internal field names:** `snake_case`, lowercase, ASCII only. Prefix with worksheet abbreviation for cross-sheet references.
- **Display names:** Title Case, human-readable, English. i18n column provides localisation key.
- **Named ranges:** `ws{N}_{InternalName}` where N = worksheet index (1-based).
- **Table names (Excel Table object):** `tbl_{WorksheetName}` with spaces replaced by underscores.

### 2.2 Data Type Definitions

| Type | Storage | Example | Constraints |
|---|---|---|---|
| Text | String, variable | `"MATH101"` | Max length enforced |
| Long Text | String, variable | `"This course introduces..."` | Max 10,000 chars |
| Integer | Number, no decimals | `120` | Whole numbers only |
| Decimal | Number, 2 decimal places | `3.50` | 0.00 – 10.00 for GPA |
| Boolean | `TRUE` / `FALSE` | `TRUE` | No 1/0, no blanks |
| Date | ISO 8601 date serial | `2026-09-01` | Excel date serial; display as `yyyy-mm-dd` |
| Time | Excel time serial | `09:00` | Display as `hh:mm` (24h) |
| Datetime | ISO 8601 serial | `2026-09-01T09:00:00` | Excel datetime serial |
| Enum | Text, constrained | `"Lecture"` | Must match defined set |
| URL | Text, hyperlink | `https://handbook.unimelb.edu.au/...` | Valid URI scheme |
| Email | Text | `"prof@uni.edu"` | Contains `@` and `.` |
| Phone | Text | `"+61 3 9999 9999"` | E.164 preferred; free-form accepted |
| Rating | Decimal, 1dp | `4.5` | 0.0 – 5.0 |
| Percentage | Decimal, 2dp | `0.85` | Stored as decimal; displayed as `85%` |
| Formula | Computed | `=SUM(...)` | No manual entry permitted |
| Currency | Decimal, 2dp | `45000.00` | Institution local currency; symbol in display format |

### 2.3 Column Property Definitions

Every column specification includes these 13 properties:

| Property | Description |
|---|---|
| Display Name | Header row text visible to users |
| Internal Name | Machine-readable field name (`snake_case`) |
| Description | Human-readable explanation of the field's purpose |
| Data Type | From the type table above |
| Max Length | For Text types; `—` for non-text types |
| Validation | Validation rule expressed in Excel data validation + formula syntax |
| Required | `Y` or `N` — whether a value must be present |
| Editable | `Y` or `N` — whether the user may modify |
| Unique | `Y` or `N` — whether values must be unique within the column |
| Sortable | `Y` or `N` — whether sorting on this column is meaningful |
| Filterable | `Y` or `N` — whether auto-filter applies |
| Searchable | `Y` or `N` — whether this column participates in global search |
| Exportable | `Y` or `N` — whether this column is included in exported artefacts |
| Visible | `Y` or `N` — whether the column is shown to the user |
| Example | Representative value |

---

## 3. Worksheet 1 — Course Overview

### 3.1 Worksheet Metadata

| Property | Value |
|---|---|
| Worksheet Name | `Course Overview` |
| Tab Color | `#1A5276` (Dark Blue) |
| Purpose | Master catalogue of all available courses across the institution |
| Description | Single source of truth for course metadata. Populated from university handbook / SIS. One row per course offering. Supports filtering by department, level, semester availability, credit system. |
| Display Order | 1 (leftmost tab) |
| Visibility | Visible |
| Excel Table Name | `tbl_Course_Overview` |
| Header Row | Row 1, frozen |
| Default Sort | Column A (`course_code`) ascending |
| Auto-filter | Enabled on all columns |

### 3.2 Column Definitions

---

**CO-01**

| Property | Value |
|---|---|
| Display Name | `Course Code` |
| Internal Name | `course_code` |
| Description | Unique institutional identifier for the course (e.g., MATH101, COMPSCI230) |
| Data Type | Text |
| Max Length | 20 |
| Validation | `=AND(LEN(A2)>=4, LEN(A2)<=20, ISTEXT(A2))` |
| Required | Y |
| Editable | N (system-populated) |
| Unique | Y |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"MATH101"` |

---

**CO-02**

| Property | Value |
|---|---|
| Display Name | `Course Name` |
| Internal Name | `course_name` |
| Description | Full title of the course |
| Data Type | Text |
| Max Length | 200 |
| Validation | `=AND(LEN(B2)>=3, LEN(B2)<=200)` |
| Required | Y |
| Editable | N |
| Unique | Y |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Calculus I"` |

---

**CO-03**

| Property | Value |
|---|---|
| Display Name | `Description` |
| Internal Name | `description` |
| Description | Full course syllabus summary; learning outcomes; topics covered |
| Data Type | Long Text |
| Max Length | 10000 |
| Validation | `=LEN(C2)<=10000` |
| Required | N |
| Editable | N |
| Unique | N |
| Sortable | N |
| Filterable | N |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"An introduction to differential and integral calculus of functions of one variable..."` |

---

**CO-04**

| Property | Value |
|---|---|
| Display Name | `Credits` |
| Internal Name | `credits` |
| Description | Numeric credit value in the institution's credit system |
| Data Type | Decimal |
| Max Length | — |
| Validation | `=AND(D2>0, D2<=240)` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `6` (US); `12.5` (AU); `15` (ECTS) |

---

**CO-05**

| Property | Value |
|---|---|
| Display Name | `Credit System` |
| Internal Name | `credit_system` |
| Description | The credit measurement framework |
| Data Type | Enum |
| Max Length | — |
| Validation | List: `US Credits`, `ECTS`, `Australian Credit Points`, `UK Credits (CATS)`, `Canadian Credits`, `Singapore Modular Credits`, `Hong Kong Credits`, `Chinese Credits`, `Other` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"US Credits"` |

---

**CO-06**

| Property | Value |
|---|---|
| Display Name | `Course Type` |
| Internal Name | `course_type` |
| Description | Classification of the course within the degree structure |
| Data Type | Enum |
| Max Length | — |
| Validation | List: see Section 10.1 |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Required"` |

---

**CO-07**

| Property | Value |
|---|---|
| Display Name | `Department` |
| Internal Name | `department` |
| Description | Academic department or school offering the course |
| Data Type | Text |
| Max Length | 100 |
| Validation | `=LEN(G2)<=100` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Department of Mathematics"` |

---

**CO-08**

| Property | Value |
|---|---|
| Display Name | `Faculty` |
| Internal Name | `faculty` |
| Description | Parent faculty or college |
| Data Type | Text |
| Max Length | 100 |
| Validation | `=LEN(H2)<=100` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Faculty of Science"` |

---

**CO-09**

| Property | Value |
|---|---|
| Display Name | `Level` |
| Internal Name | `level` |
| Description | Academic level (year of study) |
| Data Type | Enum |
| Max Length | — |
| Validation | List: `Introductory`, `Intermediate`, `Advanced`, `Honours`, `Masters`, `Doctoral` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `"Introductory"` |

---

**CO-10**

| Property | Value |
|---|---|
| Display Name | `Prerequisites` |
| Internal Name | `prerequisites` |
| Description | Semi-colon-delimited list of prerequisite course codes. Format: `MATH101; PHYS150` |
| Data Type | Long Text |
| Max Length | 2000 |
| Validation | `=LEN(J2)<=2000` |
| Required | N |
| Editable | N |
| Unique | N |
| Sortable | N |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"MATH101; MATH102"` |

---

**CO-11**

| Property | Value |
|---|---|
| Display Name | `Corequisites` |
| Internal Name | `corequisites` |
| Description | Semi-colon-delimited list of course codes that must be taken concurrently |
| Data Type | Long Text |
| Max Length | 2000 |
| Validation | `=LEN(K2)<=2000` |
| Required | N |
| Editable | N |
| Unique | N |
| Sortable | N |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"PHYS150"` |

---

**CO-12**

| Property | Value |
|---|---|
| Display Name | `Antirequisites` |
| Internal Name | `antirequisites` |
| Description | Semi-colon-delimited list of mutually exclusive courses |
| Data Type | Long Text |
| Max Length | 2000 |
| Validation | `=LEN(L2)<=2000` |
| Required | N |
| Editable | N |
| Unique | N |
| Sortable | N |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"MATH150"` |

---

**CO-13**

| Property | Value |
|---|---|
| Display Name | `Duration (Weeks)` |
| Internal Name | `duration_weeks` |
| Description | Standard teaching weeks for the course |
| Data Type | Integer |
| Max Length | — |
| Validation | `=AND(M2>=1, M2<=52, INT(M2)=M2)` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `12` |

---

**CO-14**

| Property | Value |
|---|---|
| Display Name | `Contact Hours / Week` |
| Internal Name | `contact_hours_per_week` |
| Description | Total weekly contact hours (sum of all session types) |
| Data Type | Decimal |
| Max Length | — |
| Validation | `=AND(N2>=0, N2<=60)` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `4` |

---

**CO-15**

| Property | Value |
|---|---|
| Display Name | `Assessment Types` |
| Internal Name | `assessment_types` |
| Description | Comma-separated list of assessment methods from Section 10.5 |
| Data Type | Text |
| Max Length | 500 |
| Validation | `=LEN(O2)<=500` |
| Required | N |
| Editable | N |
| Unique | N |
| Sortable | N |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Exam, Assignment, Quiz"` |

---

**CO-16**

| Property | Value |
|---|---|
| Display Name | `Enrolment Cap` |
| Internal Name | `enrolment_cap` |
| Description | Maximum number of students permitted |
| Data Type | Integer |
| Max Length | — |
| Validation | `=AND(P2>=1, P2<=5000, INT(P2)=P2)` |
| Required | N |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `300` |

---

**CO-17**

| Property | Value |
|---|---|
| Display Name | `Enrolment Difficulty` |
| Internal Name | `enrolment_difficulty` |
| Description | Historical difficulty of securing a place |
| Data Type | Enum |
| Max Length | — |
| Validation | List: see Section 10.4 |
| Required | N |
| Editable | Y (AI may override) |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `"Hard"` |

---

**CO-18**

| Property | Value |
|---|---|
| Display Name | `Student Rating` |
| Internal Name | `student_rating` |
| Description | Aggregate student satisfaction rating (0.0 – 5.0) |
| Data Type | Rating |
| Max Length | — |
| Validation | `=AND(R2>=0, R2<=5)` |
| Required | N |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `4.2` |

---

**CO-19**

| Property | Value |
|---|---|
| Display Name | `Handbook URL` |
| Internal Name | `handbook_url` |
| Description | Link to the official course page in the university handbook |
| Data Type | URL |
| Max Length | 2048 |
| Validation | `=AND(LEFT(S2,4)="http", LEN(S2)<=2048)` |
| Required | N |
| Editable | N |
| Unique | N |
| Sortable | N |
| Filterable | N |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `"https://handbook.unimelb.edu.au/2026/subjects/math101"` |

---

**CO-20**

| Property | Value |
|---|---|
| Display Name | `Available Semesters` |
| Internal Name | `available_semesters` |
| Description | Comma-separated semester codes when the course is offered (e.g., `S1,S2` or `T1,T2,T3` or `Q1,Q2,Q3,Q4`) |
| Data Type | Text |
| Max Length | 50 |
| Validation | `=LEN(T2)<=50` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | N |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"S1,S2"` |

---

**CO-21**

| Property | Value |
|---|---|
| Display Name | `i18n Key` |
| Internal Name | `i18n_key` |
| Description | Internationalisation lookup key for localised display names |
| Data Type | Text |
| Max Length | 100 |
| Validation | `=LEN(U2)<=100` |
| Required | N |
| Editable | N |
| Unique | Y |
| Sortable | N |
| Filterable | N |
| Searchable | N |
| Exportable | N |
| Visible | N |
| Example | `"course.math101"` |

---

**CO-22**

| Property | Value |
|---|---|
| Display Name | `Last Updated` |
| Internal Name | `last_updated` |
| Description | Timestamp of last modification to this record |
| Data Type | Datetime |
| Max Length | — |
| Validation | `=ISNUMBER(V2)` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `2026-07-15T10:30:00` |

---

**CO-23**

| Property | Value |
|---|---|
| Display Name | `Is Active` |
| Internal Name | `is_active` |
| Description | Whether the course is currently offered (not discontinued) |
| Data Type | Boolean |
| Max Length | — |
| Validation | `=OR(W2=TRUE, W2=FALSE)` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `TRUE` |

### 3.3 Conditional Formatting Rules — Course Overview

| Rule ID | Priority | Applies To | Condition | Format |
|---|---|---|---|---|
| CF-CO-01 | 1 | `course_type` column | `="Required"` | Dark Red fill (`#8B0000`), white text |
| CF-CO-02 | 2 | `course_type` column | `="Major Elective"` | Green fill (`#228B22`), white text |
| CF-CO-03 | 3 | `enrolment_difficulty` column | `="Very Hard"` | Red border, bold text |
| CF-CO-04 | 4 | `enrolment_difficulty` column | `="Hard"` | Orange border |
| CF-CO-05 | 5 | `student_rating` column | `>=4.5` | Gold fill (`#FFD700`) |
| CF-CO-06 | 6 | `is_active` column | `=FALSE` | Grey text, strikethrough |

### 3.4 Named Ranges — Course Overview

| Range Name | Refers To |
|---|---|
| `ws1_course_code` | `=tbl_Course_Overview[course_code]` |
| `ws1_course_name` | `=tbl_Course_Overview[course_name]` |
| `ws1_credits` | `=tbl_Course_Overview[credits]` |
| `ws1_course_type` | `=tbl_Course_Overview[course_type]` |
| `ws1_prerequisites` | `=tbl_Course_Overview[prerequisites]` |
| `ws1_available_semesters` | `=tbl_Course_Overview[available_semesters]` |

---

## 4. Worksheet 2 — Degree Planner

### 4.1 Worksheet Metadata

| Property | Value |
|---|---|
| Worksheet Name | `Degree Planner` |
| Tab Color | `#117A65` (Teal Green) |
| Purpose | Track degree requirements, course enrolment plan, completion progress, and GPA |
| Description | Two logical sections merged into one worksheet: (A) Degree Requirement Tracker — each row is a degree requirement category; (B) Semester Enrolment Plan — each row is a planned or completed course enrolment in a specific semester. Separated by a blank row and a section header row. |
| Display Order | 2 |
| Visibility | Visible |
| Excel Table Names | `tbl_Degree_Requirements` (Section A), `tbl_Enrolment_Plan` (Section B) |
| Header Row | Row 1 (frozen); Section B header at dynamic offset |

### 4.2 Section A — Degree Requirement Tracker

#### 4.2.1 Section Metadata

| Property | Value |
|---|---|
| Section Header Row | Row 1 |
| Excel Table Name | `tbl_Degree_Requirements` |
| Default Sort | `requirement_category` ascending, then `requirement_name` ascending |

#### 4.2.2 Column Definitions

---

**DR-01**

| Property | Value |
|---|---|
| Display Name | `Requirement ID` |
| Internal Name | `requirement_id` |
| Description | Unique identifier for this degree requirement line |
| Data Type | Text |
| Max Length | 30 |
| Validation | `=AND(LEN(A2)>=1, LEN(A2)<=30)` |
| Required | Y |
| Editable | N |
| Unique | Y |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"REQ-BSC-MATH-CORE"` |

---

**DR-02**

| Property | Value |
|---|---|
| Display Name | `Requirement Category` |
| Internal Name | `requirement_category` |
| Description | Broad classification: Core, Major, Elective, General Education, Capstone, Internship |
| Data Type | Text |
| Max Length | 50 |
| Validation | `=LEN(B2)<=50` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Core"` |

---

**DR-03**

| Property | Value |
|---|---|
| Display Name | `Requirement Name` |
| Internal Name | `requirement_name` |
| Description | Human-readable name of this requirement group |
| Data Type | Text |
| Max Length | 200 |
| Validation | `=LEN(C2)<=200` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Mathematics Foundation"` |

---

**DR-04**

| Property | Value |
|---|---|
| Display Name | `Min Credits Required` |
| Internal Name | `min_credits_required` |
| Description | Minimum credit points that must be earned in this category |
| Data Type | Decimal |
| Max Length | — |
| Validation | `=AND(D2>=0, D2<=500)` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `24` |

---

**DR-05**

| Property | Value |
|---|---|
| Display Name | `Credits Completed` |
| Internal Name | `credits_completed` |
| Description | Credits earned so far in this category (sum of passed courses) |
| Data Type | Formula |
| Max Length | — |
| Validation | Computed — no manual entry |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `=SUMIF(tbl_Enrolment_Plan[requirement_id], [@requirement_id], tbl_Enrolment_Plan[credits_earned])` |

---

**DR-06**

| Property | Value |
|---|---|
| Display Name | `Credits Remaining` |
| Internal Name | `credits_remaining` |
| Description | Credits still needed (min required minus completed, floor 0) |
| Data Type | Formula |
| Max Length | — |
| Validation | Computed |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `=MAX(0, [@min_credits_required] - [@credits_completed])` |

---

**DR-07**

| Property | Value |
|---|---|
| Display Name | `Status` |
| Internal Name | `status` |
| Description | Fulfilment status of this requirement |
| Data Type | Formula |
| Max Length | — |
| Validation | Computed |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `=IF([@credits_remaining]<=0, "Fulfilled", IF([@credits_completed]>0, "In Progress", "Not Started"))` |

---

**DR-08**

| Property | Value |
|---|---|
| Display Name | `Mandatory Courses` |
| Internal Name | `mandatory_courses` |
| Description | Semi-colon-delimited course codes that MUST be taken to fulfil this requirement |
| Data Type | Long Text |
| Max Length | 2000 |
| Validation | `=LEN(H2)<=2000` |
| Required | N |
| Editable | N |
| Unique | N |
| Sortable | N |
| Filterable | N |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"MATH101; MATH102; MATH201"` |

---

**DR-09**

| Property | Value |
|---|---|
| Display Name | `Notes` |
| Internal Name | `notes` |
| Description | Free-text notes about this requirement (e.g., "Must achieve minimum C grade") |
| Data Type | Long Text |
| Max Length | 1000 |
| Validation | `=LEN(I2)<=1000` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | N |
| Filterable | N |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"At least 12 credits must be at 200-level or above"` |

---

### 4.3 Section B — Semester Enrolment Plan

#### 4.3.1 Section Metadata

| Property | Value |
|---|---|
| Section Header starts | Below last row of Section A + 2 blank rows |
| Excel Table Name | `tbl_Enrolment_Plan` |
| Default Sort | `semester_code` ascending, then `priority` ascending |

#### 4.3.2 Column Definitions

---

**EP-01**

| Property | Value |
|---|---|
| Display Name | `Plan ID` |
| Internal Name | `plan_id` |
| Description | Unique identifier for this enrolment plan row |
| Data Type | Text |
| Max Length | 30 |
| Validation | `=AND(LEN(A2)>=1, LEN(A2)<=30)` |
| Required | Y |
| Editable | N |
| Unique | Y |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"PLAN-2026-S1-001"` |

---

**EP-02**

| Property | Value |
|---|---|
| Display Name | `Course Code` |
| Internal Name | `course_code` |
| Description | Foreign key to Course Overview |
| Data Type | Text |
| Max Length | 20 |
| Validation | `=COUNTIF(ws1_course_code, A2)>0` |
| Required | Y |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"MATH101"` |

---

**EP-03**

| Property | Value |
|---|---|
| Display Name | `Requirement ID` |
| Internal Name | `requirement_id` |
| Description | Foreign key to Degree Requirements; which requirement this course fulfils |
| Data Type | Text |
| Max Length | 30 |
| Validation | `=COUNTIF(tbl_Degree_Requirements[requirement_id], C2)>0` |
| Required | Y |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"REQ-BSC-MATH-CORE"` |

---

**EP-04**

| Property | Value |
|---|---|
| Display Name | `Semester Code` |
| Internal Name | `semester_code` |
| Description | Identifier for the semester/term in which the course is planned or taken |
| Data Type | Text |
| Max Length | 20 |
| Validation | `=LEN(D2)<=20` |
| Required | Y |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"2026-S1"` (semester); `"2026-T1"` (trimester); `"2026-Q1"` (quarter) |

---

**EP-05**

| Property | Value |
|---|---|
| Display Name | `Status` |
| Internal Name | `status` |
| Description | Current state of this enrolment |
| Data Type | Enum |
| Max Length | — |
| Validation | List: `Planned`, `Enrolled`, `In Progress`, `Completed`, `Withdrawn`, `Failed`, `Deferred` |
| Required | Y |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Planned"` |

---

**EP-06**

| Property | Value |
|---|---|
| Display Name | `Priority` |
| Internal Name | `priority` |
| Description | How important this course is in the enrolment sequence |
| Data Type | Enum |
| Max Length | — |
| Validation | List: see Section 10.3 |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `"Critical"` |

---

**EP-07**

| Property | Value |
|---|---|
| Display Name | `Credits Earned` |
| Internal Name | `credits_earned` |
| Description | Credits awarded upon completion (0 if not yet completed) |
| Data Type | Decimal |
| Max Length | — |
| Validation | `=OR(G2=0, AND(G2>0, G2<=240))` |
| Required | Y |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `6` |

---

**EP-08**

| Property | Value |
|---|---|
| Display Name | `Grade` |
| Internal Name | `grade` |
| Description | Letter or numeric grade received |
| Data Type | Text |
| Max Length | 10 |
| Validation | `=LEN(H2)<=10` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"A"` / `"H1"` / `"85"` |

---

**EP-09**

| Property | Value |
|---|---|
| Display Name | `Grade Points` |
| Internal Name | `grade_points` |
| Description | Numeric grade point (e.g., 4.0 for A, 3.0 for B in US GPA) |
| Data Type | Decimal |
| Max Length | — |
| Validation | `=AND(I2>=0, I2<=10)` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `4.0` |

---

**EP-10**

| Property | Value |
|---|---|
| Display Name | `Notes` |
| Internal Name | `notes` |
| Description | Free-text notes (e.g., "Need override from department") |
| Data Type | Long Text |
| Max Length | 1000 |
| Validation | `=LEN(J2)<=1000` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | N |
| Filterable | N |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Prerequisite waiver granted 2026-07-01"` |

---

### 4.4 Summary Row — Degree Planner

A summary row is placed at the bottom of the Degree Planner worksheet, below Section B, with the following computed values:

| Metric | Formula |
|---|---|
| Total Credits Required | `=SUM(tbl_Degree_Requirements[min_credits_required])` |
| Total Credits Completed | `=SUMIF(tbl_Enrolment_Plan[status], "Completed", tbl_Enrolment_Plan[credits_earned])` |
| Total Credits Remaining | `=Total_Credits_Required - Total_Credits_Completed` |
| Graduation Progress | `=Total_Credits_Completed / Total_Credits_Required` (display as %) |
| Cumulative GPA | `=SUMPRODUCT(tbl_Enrolment_Plan[credits_earned], tbl_Enrolment_Plan[grade_points]) / SUMIF(tbl_Enrolment_Plan[credits_earned], ">0", tbl_Enrolment_Plan[credits_earned])` |
| Current Semester Credits | `=SUMIFS(tbl_Enrolment_Plan[credits_earned], tbl_Enrolment_Plan[semester_code], Current_Semester_Code)` |
| Current Semester GPA | `=SUMPRODUCT(...)` with same logic filtered to current semester |

### 4.5 Conditional Formatting Rules — Degree Planner

| Rule ID | Priority | Applies To | Condition | Format |
|---|---|---|---|---|
| CF-DP-01 | 1 | `status` (Section A) | `="Fulfilled"` | Green fill (`#27AE60`), white text |
| CF-DP-02 | 2 | `status` (Section A) | `="Not Started"` | Light grey fill (`#D5D8DC`) |
| CF-DP-03 | 3 | `credits_remaining` | `>0` | Yellow fill (`#F1C40F`) |
| CF-DP-04 | 4 | `status` (Section B) | `="Failed"` | Red fill (`#E74C3C`), white text, bold |
| CF-DP-05 | 5 | `status` (Section B) | `="Completed"` | Green fill (`#27AE60`), white text |
| CF-DP-06 | 6 | `status` (Section B) | `="Planned"` | Blue fill (`#3498DB`), white text |
| CF-DP-07 | 7 | `priority` column | `="Critical"` | Dark red text, bold |
| CF-DP-08 | 8 | Graduation Progress cell | `< 0.5` | Red fill |
| CF-DP-09 | 9 | Graduation Progress cell | `>= 1` | Gold fill (`#FFD700`), bold |
| CF-DP-10 | 10 | Current Semester Credits cell | `> 20` | Red border (credit overload warning) |

---

## 5. Worksheet 3 — AI Recommendations

### 5.1 Worksheet Metadata

| Property | Value |
|---|---|
| Worksheet Name | `AI Recommendations` |
| Tab Color | `#1ABC9C` (Cyan) |
| Purpose | AI-generated course suggestions with reasoning, confidence scores, and alternatives |
| Description | Each row is a recommendation. AI populates this sheet. Users may accept (moving row to Enrolment Plan), reject, or snooze. Includes audit trail. |
| Display Order | 3 |
| Visibility | Visible |
| Excel Table Name | `tbl_AI_Recommendations` |
| Header Row | Row 1, frozen |

### 5.2 Column Definitions

---

**AI-01**

| Property | Value |
|---|---|
| Display Name | `Recommendation ID` |
| Internal Name | `recommendation_id` |
| Description | Unique identifier for this recommendation |
| Data Type | Text |
| Max Length | 30 |
| Validation | `=AND(LEN(A2)>=1, LEN(A2)<=30)` |
| Required | Y |
| Editable | N |
| Unique | Y |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"REC-20260803-0001"` |

---

**AI-02**

| Property | Value |
|---|---|
| Display Name | `Course Code` |
| Internal Name | `course_code` |
| Description | Recommended course (FK to Course Overview) |
| Data Type | Text |
| Max Length | 20 |
| Validation | `=COUNTIF(ws1_course_code, B2)>0` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"MATH201"` |

---

**AI-03**

| Property | Value |
|---|---|
| Display Name | `Reason` |
| Internal Name | `reason` |
| Description | Natural language explanation of why this course is recommended |
| Data Type | Long Text |
| Max Length | 2000 |
| Validation | `=LEN(C2)<=2000` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | N |
| Filterable | N |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Fulfils remaining Core Maths requirement (REQ-BSC-MATH-CORE). Prerequisites satisfied. High historical student satisfaction."` |

---

**AI-04**

| Property | Value |
|---|---|
| Display Name | `Recommendation Type` |
| Internal Name | `recommendation_type` |
| Description | Category of recommendation logic |
| Data Type | Enum |
| Max Length | — |
| Validation | List: `Requirement Fulfilment`, `Interest Match`, `Career Path`, `Prerequisite Chain`, `Workload Balance`, `Popular Course`, `Peer Recommendation`, `Academic Advisor` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `"Requirement Fulfilment"` |

---

**AI-05**

| Property | Value |
|---|---|
| Display Name | `Confidence Score` |
| Internal Name | `confidence_score` |
| Description | AI confidence in this recommendation (0.00 – 1.00) |
| Data Type | Percentage |
| Max Length | — |
| Validation | `=AND(E2>=0, E2<=1)` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `0.92` |

---

**AI-06**

| Property | Value |
|---|---|
| Display Name | `Priority` |
| Internal Name | `priority` |
| Description | Urgency of this recommendation |
| Data Type | Enum |
| Max Length | — |
| Validation | List: see Section 10.3 |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `"High"` |

---

**AI-07**

| Property | Value |
|---|---|
| Display Name | `Recommended Semester` |
| Internal Name | `recommended_semester` |
| Description | Which semester/term the AI suggests enrolling |
| Data Type | Text |
| Max Length | 20 |
| Validation | `=LEN(G2)<=20` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"2026-S2"` |

---

**AI-08**

| Property | Value |
|---|---|
| Display Name | `Alternatives` |
| Internal Name | `alternatives` |
| Description | Semi-colon-delimited list of alternative course codes |
| Data Type | Long Text |
| Max Length | 2000 |
| Validation | `=LEN(H2)<=2000` |
| Required | N |
| Editable | N |
| Unique | N |
| Sortable | N |
| Filterable | N |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"MATH202; PHYS201; COMPSCI230"` |

---

**AI-09**

| Property | Value |
|---|---|
| Display Name | `Decision` |
| Internal Name | `decision` |
| Description | User's response to this recommendation |
| Data Type | Enum |
| Max Length | — |
| Validation | List: `Pending`, `Accepted`, `Rejected`, `Snoozed` |
| Required | Y |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `"Pending"` |

---

**AI-10**

| Property | Value |
|---|---|
| Display Name | `Decision Timestamp` |
| Internal Name | `decision_timestamp` |
| Description | When the user made their decision |
| Data Type | Datetime |
| Max Length | — |
| Validation | `=ISNUMBER(J2)` |
| Required | N |
| Editable | N (system-set) |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `2026-08-03T14:30:00` |

---

**AI-11**

| Property | Value |
|---|---|
| Display Name | `Resulting Plan ID` |
| Internal Name | `resulting_plan_id` |
| Description | If accepted, the Plan ID created in Enrolment Plan (FK) |
| Data Type | Text |
| Max Length | 30 |
| Validation | `=OR(K2="", COUNTIF(tbl_Enrolment_Plan[plan_id], K2)>0)` |
| Required | N |
| Editable | N |
| Unique | N |
| Sortable | N |
| Filterable | N |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"PLAN-2026-S2-005"` |

---

**AI-12**

| Property | Value |
|---|---|
| Display Name | `Generated At` |
| Internal Name | `generated_at` |
| Description | When AI created this recommendation |
| Data Type | Datetime |
| Max Length | — |
| Validation | `=ISNUMBER(L2)` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `2026-08-03T09:00:00` |

---

**AI-13**

| Property | Value |
|---|---|
| Display Name | `Expires At` |
| Internal Name | `expires_at` |
| Description | When this recommendation becomes stale (e.g., enrolment deadline passed) |
| Data Type | Datetime |
| Max Length | — |
| Validation | `=AND(ISNUMBER(M2), M2 > [@generated_at])` |
| Required | N |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `2026-12-31T23:59:59` |

### 5.3 Conditional Formatting Rules — AI Recommendations

| Rule ID | Priority | Applies To | Condition | Format |
|---|---|---|---|---|
| CF-AI-01 | 1 | Entire row | `decision="Accepted"` | Green fill (`#27AE60`), white text |
| CF-AI-02 | 2 | Entire row | `decision="Rejected"` | Grey fill, strikethrough |
| CF-AI-03 | 3 | `confidence_score` column | `>=0.9` | Gold fill (`#FFD700`) |
| CF-AI-04 | 4 | `confidence_score` column | `<0.5` | Light red fill (`#FADBD8`) |
| CF-AI-05 | 5 | `expires_at` column | `<TODAY()` | Red text, bold |

---

## 6. Worksheet 4 — Weekly Timetable

### 6.1 Worksheet Metadata

| Property | Value |
|---|---|
| Worksheet Name | `Weekly Timetable` |
| Tab Color | `#8E44AD` (Purple) |
| Purpose | Visual weekly grid showing class sessions per day and time slot |
| Description | Grid layout: Day columns (Monday–Friday) with optional Saturday/Sunday; Time rows (08:00–22:00 in 30-min increments). Cells are merged where sessions span multiple rows. Each occupied cell displays course code, session type, location, instructor. |
| Display Order | 4 |
| Visibility | Visible |
| Header Row | Row 1 (day headers), Column A (time labels), both frozen |
| Grid Range | `B2:H57` (28 time slots × 7 days default; adjustable) |

### 6.2 Timetable Grid Layout

#### 6.2.1 Structure

| Element | Value |
|---|---|
| Row 1 | Day headers: `Time`, `Monday`, `Tuesday`, `Wednesday`, `Thursday`, `Friday`, `Saturday`, `Sunday` |
| Column A | Time labels: `08:00`, `08:30`, `09:00`, ... , `21:30` (30-minute blocks) |
| Data area | `B2:H57` |
| Daily hour count | `=ROWS(Data_Area) * 0.5` hours per day |
| Week pattern selector | Cell `A1` or separate control: dropdown `Every Week`, `Odd Weeks`, `Even Weeks`, `Custom` |

#### 6.2.2 Cell Content Format

Each occupied grid cell displays:

```
{course_code}
{session_type}
{location}
{instructor_last_name}
```

Cells for the same course × session type are merged vertically (via Excel merge or rendered via conditional formatting with borders).

#### 6.2.3 Conditional Formatting Rules — Weekly Timetable

| Rule ID | Priority | Applies To | Condition (checks cell text start) | Format |
|---|---|---|---|---|
| CF-TT-01 | 1 | Entire grid | Cell contains course code AND course is `Required` (via VLOOKUP to Course Overview) | Dark Red fill (`#8B0000`), white text |
| CF-TT-02 | 2 | Entire grid | Cell contains course code AND course is `Major Elective` | Green fill (`#228B22`), white text |
| CF-TT-03 | 3 | Entire grid | Cell contains course code AND course is `General Education` | Blue fill (`#2980B9`), white text |
| CF-TT-04 | 4 | Entire grid | Cell contains `"Lab"` | Purple fill (`#8E44AD`), white text |
| CF-TT-05 | 5 | Entire grid | Cell contains `"Tutorial"` | Light Blue fill (`#85C1E9`), dark text |
| CF-TT-06 | 6 | Entire grid | Cell contains `"Workshop"` | Orange fill (`#E67E22`), white text |
| CF-TT-07 | 7 | Entire grid | Cell contains `"Exam"` | Red border (3pt), bold |
| CF-TT-08 | 8 | Entire grid | Cell is blank | Grey fill (`#EAECEE`) — Free Time |
| CF-TT-09 | 9 | Entire grid | Time overlap detected (two courses in same time block) | Red fill (`#FF0000`), yellow text, bold — Conflict Warning |

#### 6.2.4 Time Overlap Detection

| Property | Value |
|---|---|
| Detection method | COMPARE each cell against Raw Schedule Database for same day + overlapping time range |
| Visual indicator | Red cell fill + `⚠ CONFLICT` text appended |
| Formula concept | For each grid cell at day D, time T: `=IF(COUNTIFS(RawSchedule[day], D, RawSchedule[start_time], "<="&T, RawSchedule[end_time], ">"&T) > 1, "CONFLICT", "")` |

#### 6.2.5 Week Pattern Support

| Pattern | Implementation |
|---|---|
| Every Week | Default; all sessions rendered |
| Odd Weeks | Filter Raw Schedule Database for odd week numbers, render only those |
| Even Weeks | Filter Raw Schedule Database for even week numbers, render only those |
| Week Range (e.g., Weeks 1–6) | Filter by `start_week` <= target week <= `end_week` |
| Single-Session Exception | Filter by `excluded_dates` from Raw Schedule Database |

---

## 7. Worksheet 5 — Academic Calendar

### 7.1 Worksheet Metadata

| Property | Value |
|---|---|
| Worksheet Name | `Academic Calendar` |
| Tab Color | `#E67E22` (Orange) |
| Purpose | Institution-wide academic dates: semesters, breaks, exams, deadlines |
| Description | Calendar of all significant academic dates. Supports export to personal calendar (.ics, Google, Outlook). One row per event. |
| Display Order | 5 |
| Visibility | Visible |
| Excel Table Name | `tbl_Academic_Calendar` |
| Header Row | Row 1, frozen |
| Default Sort | `start_date` ascending |

### 7.2 Column Definitions

---

**AC-01**

| Property | Value |
|---|---|
| Display Name | `Event ID` |
| Internal Name | `event_id` |
| Description | Unique event identifier |
| Data Type | Text |
| Max Length | 30 |
| Validation | `=AND(LEN(A2)>=1, LEN(A2)<=30)` |
| Required | Y |
| Editable | N |
| Unique | Y |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"EVT-2026-ORIENTATION"` |

---

**AC-02**

| Property | Value |
|---|---|
| Display Name | `Event Name` |
| Internal Name | `event_name` |
| Description | Human-readable event title |
| Data Type | Text |
| Max Length | 200 |
| Validation | `=AND(LEN(B2)>=1, LEN(B2)<=200)` |
| Required | Y |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Semester 1 Orientation Week"` |

---

**AC-03**

| Property | Value |
|---|---|
| Display Name | `Event Type` |
| Internal Name | `event_type` |
| Description | Classification of the calendar event |
| Data Type | Enum |
| Max Length | — |
| Validation | List: `Semester Start`, `Semester End`, `Teaching Start`, `Teaching End`, `Exam Period`, `Break`, `Holiday`, `Census Date`, `Add/Drop Deadline`, `Withdrawal Deadline`, `Results Release`, `Orientation`, `Graduation`, `Other` |
| Required | Y |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Orientation"` |

---

**AC-04**

| Property | Value |
|---|---|
| Display Name | `Start Date` |
| Internal Name | `start_date` |
| Description | Event start date |
| Data Type | Date |
| Max Length | — |
| Validation | `=AND(ISNUMBER(D2), D2>=DATE(2020,1,1))` |
| Required | Y |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `2026-02-23` |

---

**AC-05**

| Property | Value |
|---|---|
| Display Name | `End Date` |
| Internal Name | `end_date` |
| Description | Event end date (same as start for single-day events) |
| Data Type | Date |
| Max Length | — |
| Validation | `=AND(ISNUMBER(E2), E2>=[@start_date])` |
| Required | Y |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `2026-02-27` |

---

**AC-06**

| Property | Value |
|---|---|
| Display Name | `Start Time` |
| Internal Name | `start_time` |
| Description | Event start time (for timed events; blank for all-day events) |
| Data Type | Time |
| Max Length | — |
| Validation | `=OR(F2="", AND(F2>=TIME(0,0,0), F2<=TIME(23,59,0)))` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `09:00` |

---

**AC-07**

| Property | Value |
|---|---|
| Display Name | `End Time` |
| Internal Name | `end_time` |
| Description | Event end time |
| Data Type | Time |
| Max Length | — |
| Validation | `=OR(G2="", G2>=[@start_time])` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `17:00` |

---

**AC-08**

| Property | Value |
|---|---|
| Display Name | `Semester Code` |
| Internal Name | `semester_code` |
| Description | Which semester/term this event belongs to |
| Data Type | Text |
| Max Length | 20 |
| Validation | `=LEN(H2)<=20` |
| Required | Y |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"2026-S1"` |

---

**AC-09**

| Property | Value |
|---|---|
| Display Name | `Is Recurring` |
| Internal Name | `is_recurring` |
| Description | Whether this event repeats |
| Data Type | Boolean |
| Max Length | — |
| Validation | `=OR(I2=TRUE, I2=FALSE)` |
| Required | Y |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `TRUE` |

---

**AC-10**

| Property | Value |
|---|---|
| Display Name | `Recurrence Rule` |
| Internal Name | `recurrence_rule` |
| Description | RFC 5545 RRULE string (e.g., `FREQ=WEEKLY;BYDAY=MO,WE;COUNT=12`) |
| Data Type | Text |
| Max Length | 500 |
| Validation | `=OR(J2="", LEFT(J2,5)="FREQ=")` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | N |
| Filterable | N |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `"FREQ=WEEKLY;BYDAY=MO,WE,FR;COUNT=12"` |

---

**AC-11**

| Property | Value |
|---|---|
| Display Name | `Location` |
| Internal Name | `location` |
| Description | Physical or virtual location |
| Data Type | Text |
| Max Length | 200 |
| Validation | `=LEN(K2)<=200` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | N |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Main Auditorium"` |

---

**AC-12**

| Property | Value |
|---|---|
| Display Name | `Description` |
| Internal Name | `description` |
| Description | Detailed event information |
| Data Type | Long Text |
| Max Length | 5000 |
| Validation | `=LEN(L2)<=5000` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | N |
| Filterable | N |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Mandatory attendance for all first-year Science students..."` |

---

**AC-13**

| Property | Value |
|---|---|
| Display Name | `Calendar UID` |
| Internal Name | `calendar_uid` |
| Description | Globally unique identifier for calendar sync (RFC 5545 UID) |
| Data Type | Text |
| Max Length | 255 |
| Validation | `=LEN(M2)<=255` |
| Required | N |
| Editable | N |
| Unique | Y |
| Sortable | N |
| Filterable | N |
| Searchable | N |
| Exportable | Y |
| Visible | N |
| Example | `"a1b2c3d4-e5f6-7890-abcd-ef1234567890@uni.edu"` |

---

**AC-14**

| Property | Value |
|---|---|
| Display Name | `Reminder (Minutes)` |
| Internal Name | `reminder_minutes` |
| Description | Minutes before event to trigger reminder (0 = no reminder) |
| Data Type | Integer |
| Max Length | — |
| Validation | `=AND(N2>=0, N2<=10080, INT(N2)=N2)` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `60` |

### 7.3 Conditional Formatting Rules — Academic Calendar

| Rule ID | Priority | Applies To | Condition | Format |
|---|---|---|---|---|
| CF-AC-01 | 1 | Entire row | `event_type="Holiday"` | Gold fill (`#F1C40F`) |
| CF-AC-02 | 2 | Entire row | `event_type="Exam Period"` | Red fill (`#E74C3C`), white text |
| CF-AC-03 | 3 | Entire row | `event_type="Break"` | Light blue fill (`#D6EAF8`) |
| CF-AC-04 | 4 | Entire row | `event_type` contains `"Deadline"` | Orange border, bold |
| CF-AC-05 | 5 | `end_date` column | `<TODAY()` | Grey text |
| CF-AC-06 | 6 | `start_date` column | `=TODAY()` | Green fill, bold |

---

## 8. Worksheet 6 — Raw Schedule Database

### 8.1 Worksheet Metadata

| Property | Value |
|---|---|
| Worksheet Name | `Raw Schedule Database` |
| Tab Color | `#7F8C8D` (Grey) |
| Purpose | Fully normalised database of every individual class session. The single source of truth from which the Weekly Timetable grid and all calendar exports are generated. |
| Description | One row per distinct session. A course with a lecture on Monday 09:00–10:00 and a tutorial on Wednesday 14:00–15:00 has two rows. Supports one-off sessions, recurring patterns, multiple instructors, and venue changes. |
| Display Order | 6 (rightmost tab) |
| Visibility | Hidden by default (may be made visible for power users) |
| Excel Table Name | `tbl_Raw_Schedule` |
| Header Row | Row 1, frozen |
| Default Sort | `day` (custom sort: Mon–Sun), then `start_time` ascending |

### 8.2 Day Sort Order (Custom List)

`Monday`, `Tuesday`, `Wednesday`, `Thursday`, `Friday`, `Saturday`, `Sunday`

### 8.3 Column Definitions

---

**RS-01**

| Property | Value |
|---|---|
| Display Name | `Session ID` |
| Internal Name | `session_id` |
| Description | Unique session identifier |
| Data Type | Text |
| Max Length | 30 |
| Validation | `=AND(LEN(A2)>=1, LEN(A2)<=30)` |
| Required | Y |
| Editable | N |
| Unique | Y |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"SESS-MATH101-LEC-01"` |

---

**RS-02**

| Property | Value |
|---|---|
| Display Name | `Course Code` |
| Internal Name | `course_code` |
| Description | FK to Course Overview |
| Data Type | Text |
| Max Length | 20 |
| Validation | `=COUNTIF(ws1_course_code, B2)>0` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"MATH101"` |

---

**RS-03**

| Property | Value |
|---|---|
| Display Name | `Session Type` |
| Internal Name | `session_type` |
| Description | Type of class session |
| Data Type | Enum |
| Max Length | — |
| Validation | List: see Section 10.2 |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Lecture"` |

---

**RS-04**

| Property | Value |
|---|---|
| Display Name | `Day` |
| Internal Name | `day` |
| Description | Day of the week |
| Data Type | Enum |
| Max Length | — |
| Validation | List: `Monday`, `Tuesday`, `Wednesday`, `Thursday`, `Friday`, `Saturday`, `Sunday` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y (custom sort) |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Monday"` |

---

**RS-05**

| Property | Value |
|---|---|
| Display Name | `Start Time` |
| Internal Name | `start_time` |
| Description | Session start time (24h) |
| Data Type | Time |
| Max Length | — |
| Validation | `=AND(E2>=TIME(6,0,0), E2<=TIME(22,0,0))` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `09:00` |

---

**RS-06**

| Property | Value |
|---|---|
| Display Name | `End Time` |
| Internal Name | `end_time` |
| Description | Session end time (24h); must be > start_time |
| Data Type | Time |
| Max Length | — |
| Validation | `=AND(F2>=[@start_time], F2<=TIME(22,0,0))` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `10:00` |

---

**RS-07**

| Property | Value |
|---|---|
| Display Name | `Duration (Minutes)` |
| Internal Name | `duration_minutes` |
| Description | Computed session duration |
| Data Type | Formula |
| Max Length | — |
| Validation | Computed |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `=(F2-E2)*1440` |

---

**RS-08**

| Property | Value |
|---|---|
| Display Name | `Location` |
| Internal Name | `location` |
| Description | Room, building, or virtual meeting link |
| Data Type | Text |
| Max Length | 200 |
| Validation | `=LEN(H2)<=200` |
| Required | Y |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Old Arts 101"` |

---

**RS-09**

| Property | Value |
|---|---|
| Display Name | `Instructor` |
| Internal Name | `instructor` |
| Description | Primary instructor name |
| Data Type | Text |
| Max Length | 150 |
| Validation | `=LEN(I2)<=150` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Dr. Jane Smith"` |

---

**RS-10**

| Property | Value |
|---|---|
| Display Name | `Instructor Email` |
| Internal Name | `instructor_email` |
| Description | Instructor contact email |
| Data Type | Email |
| Max Length | 254 |
| Validation | `=AND(ISNUMBER(SEARCH("@",J2)), ISNUMBER(SEARCH(".",J2)))` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"j.smith@unimelb.edu.au"` |

---

**RS-11**

| Property | Value |
|---|---|
| Display Name | `Additional Instructors` |
| Internal Name | `additional_instructors` |
| Description | Semi-colon-delimited list of co-instructors or tutors |
| Data Type | Text |
| Max Length | 500 |
| Validation | `=LEN(K2)<=500` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | N |
| Filterable | N |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Dr. Alan Turing; Prof. Ada Lovelace"` |

---

**RS-12**

| Property | Value |
|---|---|
| Display Name | `Week Pattern` |
| Internal Name | `week_pattern` |
| Description | Recurrence pattern for this session across teaching weeks |
| Data Type | Enum |
| Max Length | — |
| Validation | List: `Every Week`, `Odd Weeks`, `Even Weeks`, `Custom` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `"Every Week"` |

---

**RS-13**

| Property | Value |
|---|---|
| Display Name | `Start Week` |
| Internal Name | `start_week` |
| Description | First teaching week number this session runs |
| Data Type | Integer |
| Max Length | — |
| Validation | `=AND(M2>=1, M2<=52, INT(M2)=M2)` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `1` |

---

**RS-14**

| Property | Value |
|---|---|
| Display Name | `End Week` |
| Internal Name | `end_week` |
| Description | Last teaching week number this session runs |
| Data Type | Integer |
| Max Length | — |
| Validation | `=AND(N2>=[@start_week], N2<=52, INT(N2)=N2)` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `12` |

---

**RS-15**

| Property | Value |
|---|---|
| Display Name | `Custom Weeks` |
| Internal Name | `custom_weeks` |
| Description | Comma-separated week numbers when `week_pattern="Custom"` (e.g., `2,3,5,7,8`) |
| Data Type | Text |
| Max Length | 200 |
| Validation | `=OR(O2="", AND([@week_pattern]="Custom", LEN(O2)>0))` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | N |
| Filterable | N |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `"1,3,5,7,9,11"` |

---

**RS-16**

| Property | Value |
|---|---|
| Display Name | `Excluded Dates` |
| Internal Name | `excluded_dates` |
| Description | Comma-separated ISO dates when this session is cancelled (e.g., public holidays) |
| Data Type | Text |
| Max Length | 500 |
| Validation | `=LEN(P2)<=500` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | N |
| Filterable | N |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `"2026-03-14, 2026-04-25"` |

---

**RS-17**

| Property | Value |
|---|---|
| Display Name | `Semester Code` |
| Internal Name | `semester_code` |
| Description | Which semester/term this session belongs to |
| Data Type | Text |
| Max Length | 20 |
| Validation | `=LEN(Q2)<=20` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"2026-S1"` |

---

**RS-18**

| Property | Value |
|---|---|
| Display Name | `Date Range Start` |
| Internal Name | `date_range_start` |
| Description | First calendar date of this session instance (computed from semester start + week number + day) |
| Data Type | Date |
| Max Length | — |
| Validation | `=ISNUMBER(R2)` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `2026-03-02` |

---

**RS-19**

| Property | Value |
|---|---|
| Display Name | `Date Range End` |
| Internal Name | `date_range_end` |
| Description | Last calendar date of this session |
| Data Type | Date |
| Max Length | — |
| Validation | `=AND(ISNUMBER(S2), S2>=[@date_range_start])` |
| Required | Y |
| Editable | N |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `2026-05-25` |

---

**RS-20**

| Property | Value |
|---|---|
| Display Name | `Max Capacity` |
| Internal Name | `max_capacity` |
| Description | Room or session capacity |
| Data Type | Integer |
| Max Length | — |
| Validation | `=AND(T2>=1, T2<=2000, INT(T2)=T2)` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | Y |
| Filterable | Y |
| Searchable | N |
| Exportable | Y |
| Visible | Y |
| Example | `200` |

---

**RS-21**

| Property | Value |
|---|---|
| Display Name | `Notes` |
| Internal Name | `notes` |
| Description | Session-specific notes (e.g., "Weeks 6 & 12 in Computer Lab B instead") |
| Data Type | Long Text |
| Max Length | 1000 |
| Validation | `=LEN(U2)<=1000` |
| Required | N |
| Editable | Y |
| Unique | N |
| Sortable | N |
| Filterable | N |
| Searchable | Y |
| Exportable | Y |
| Visible | Y |
| Example | `"Weeks 6 and 12 relocated to Redmond Barry 201"` |

### 8.4 Conflict Detection Formulas

Stored as a hidden helper column (`conflict_flag`):

```
=IF(
  COUNTIFS(
    tbl_Raw_Schedule[course_code], "<>" & [@course_code],
    tbl_Raw_Schedule[day], [@day],
    tbl_Raw_Schedule[start_time], "<" & [@end_time],
    tbl_Raw_Schedule[end_time], ">" & [@start_time],
    tbl_Raw_Schedule[semester_code], [@semester_code]
  ) > 0,
  TRUE,
  FALSE
)
```

### 8.5 Conditional Formatting Rules — Raw Schedule Database

| Rule ID | Priority | Applies To | Condition | Format |
|---|---|---|---|---|
| CF-RS-01 | 1 | Entire row | `conflict_flag=TRUE` | Red fill (`#FF0000`), yellow text, bold — `⚠ TIMETABLE CONFLICT` |
| CF-RS-02 | 2 | `session_type` column | `="Lecture"` | Dark red fill (`#8B0000`), white text |
| CF-RS-03 | 3 | `session_type` column | `="Lab"` | Purple fill (`#8E44AD`), white text |
| CF-RS-04 | 4 | `session_type` column | `="Tutorial"` | Blue fill (`#2980B9`), white text |
| CF-RS-05 | 5 | `session_type` column | `="Workshop"` | Orange fill (`#E67E22`), white text |
| CF-RS-06 | 6 | `session_type` column | `="Seminar"` | Teal fill (`#1ABC9C`), white text |
| CF-RS-07 | 7 | `session_type` column | `="Exam"` | Red border (3pt), bold |
| CF-RS-08 | 8 | `week_pattern` column | `="Odd Weeks"` | Light yellow fill |
| CF-RS-09 | 9 | `week_pattern` column | `="Even Weeks"` | Light green fill |

---

## 9. Cross-Worksheet Relationships

### 9.1 Entity Relationship Diagram (Logical)

```
Course Overview (1) ────< (N) Raw Schedule Database
      │                              │
      │                              │
      │ (N)                          │ (N)
      ▼                              ▼
Degree Planner ────< Enrolment Plan      Weekly Timetable (rendered from)
      │                   │
      │                   │ (1:0..1)
      ▼                   ▼
Summary Metrics      AI Recommendations
                           │
                           │ (1:0..1)
                           ▼
                      Enrolment Plan (when accepted)

Academic Calendar is standalone but references semester_code.
```

### 9.2 Foreign Key Map

| Child Table | Child Column | Parent Table | Parent Column | Cardinality |
|---|---|---|---|---|
| `tbl_Enrolment_Plan` | `course_code` | `tbl_Course_Overview` | `course_code` | N:1 |
| `tbl_Enrolment_Plan` | `requirement_id` | `tbl_Degree_Requirements` | `requirement_id` | N:1 |
| `tbl_AI_Recommendations` | `course_code` | `tbl_Course_Overview` | `course_code` | N:1 |
| `tbl_AI_Recommendations` | `resulting_plan_id` | `tbl_Enrolment_Plan` | `plan_id` | 0..1:1 |
| `tbl_Raw_Schedule` | `course_code` | `tbl_Course_Overview` | `course_code` | N:1 |

### 9.3 Cross-Worksheet Named Ranges for Formulas

| Range Name | Refers To | Used In |
|---|---|---|
| `ws1_course_code` | `=tbl_Course_Overview[course_code]` | All FK validations |
| `ws1_course_type` | `=tbl_Course_Overview[course_type]` | Timetable colouring |
| `ws1_credits` | `=tbl_Course_Overview[credits]` | Enrolment Plan credit lookup |
| `ws2_plan_credits` | `=tbl_Enrolment_Plan[credits_earned]` | Degree requirement summary |
| `ws2_plan_status` | `=tbl_Enrolment_Plan[status]` | Progress calculations |
| `ws2_plan_semester` | `=tbl_Enrolment_Plan[semester_code]` | Current semester filtering |
| `ws6_day` | `=tbl_Raw_Schedule[day]` | Timetable rendering |
| `ws6_start_time` | `=tbl_Raw_Schedule[start_time]` | Overlap detection |
| `ws6_end_time` | `=tbl_Raw_Schedule[end_time]` | Overlap detection |

---

## 10. Enum Definitions

### 10.1 Course Type

| Value | Description |
|---|---|
| `Required` | Mandatory for the degree; must be passed |
| `Major Elective` | Elective within the major; choose from a constrained list |
| `General Education` | Breadth / general education requirement |
| `Free Elective` | Any course from any faculty; no constraints except prerequisites |

### 10.2 Session Type

| Value | Typical Duration | Icon (for UI) |
|---|---|---|
| `Lecture` | 60–120 min | 📖 |
| `Tutorial` | 60 min | ✏️ |
| `Lab` | 120–180 min | 🔬 |
| `Workshop` | 120–180 min | 🛠️ |
| `Seminar` | 60–120 min | 💬 |
| `Exam` | 120–180 min | 📝 |

### 10.3 Priority

| Value | RGB | Description |
|---|---|---|
| `Critical` | `#C0392B` | Must-enrol; failing to secure this course prevents graduation on time |
| `High` | `#E74C3C` | Strongly recommended; limited availability or prerequisite chain |
| `Medium` | `#F39C12` | Standard recommendation |
| `Low` | `#27AE60` | Nice-to-have; can defer to later semester |

### 10.4 Enrolment Difficulty

| Value | Description |
|---|---|
| `Very Hard` | Historically fills within minutes of enrolment opening; waitlist common |
| `Hard` | Usually fills within the first day; plan backup courses |
| `Normal` | Typically has spaces throughout enrolment period |
| `Easy` | Rarely reaches capacity; safe to enrol late |

### 10.5 Assessment Type

| Value |
|---|
| `Exam` |
| `Quiz` |
| `Assignment` |
| `Project` |
| `Presentation` |
| `Participation` |
| `Lab Report` |
| `Portfolio` |
| `Thesis` |
| `Practical` |

### 10.6 Credit System

| Value | Typical Full-Time Semester Load | Typical Total Degree |
|---|---|---|
| `US Credits` | 12–18 | 120–130 |
| `ECTS` | 30 | 180–240 (bachelor) |
| `Australian Credit Points` | 24–30 | 144–192 (bachelor; varies by institution, e.g., UniMelb 300) |
| `UK Credits (CATS)` | 60 | 360 (bachelor) |
| `Canadian Credits` | 15 | 120 |
| `Singapore Modular Credits` | 20 | 160 |
| `Hong Kong Credits` | 15–18 | 120 |
| `Chinese Credits` | 20–25 | 140–170 |
| `Other` | Varies | Varies |

---

## 11. Color Rules (Master Palette)

### 11.1 Course & Session Colours

| Rule | Colour Name | Hex Code | RGB | Applies To |
|---|---|---|---|---|
| Required Course | Dark Red | `#8B0000` | 139, 0, 0 | `course_type="Required"` — cell fill; text white |
| Popular Course | Orange | `#E67E22` | 230, 126, 34 | Highlighted courses with `student_rating >= 4.5` AND `enrolment_difficulty IN ("Very Hard", "Hard")` |
| Elective | Green | `#228B22` | 34, 139, 34 | `course_type="Major Elective"` or `"Free Elective"` — cell fill; text white |
| Lab | Purple | `#8E44AD` | 142, 68, 173 | `session_type="Lab"` — cell fill; text white |
| Tutorial | Blue | `#2980B9` | 41, 128, 185 | `session_type="Tutorial"` — cell fill; text white |
| AI Recommendation | Cyan | `#1ABC9C` | 26, 188, 156 | AI Recommendations sheet — tab colour and accepted row highlight |
| Study Block | Yellow | `#F1C40F` | 241, 196, 15 | User-defined study blocks in timetable |
| Free Time | Grey | `#EAECEE` | 234, 236, 238 | Unoccupied timetable cells |
| Warning | Red Border | `#FF0000` | 255, 0, 0 | `conflict_flag=TRUE`, credit overload, missing prerequisites — 3pt red border |

### 11.2 Status Colours

| Status | Colour Name | Hex Code | Applies To |
|---|---|---|---|
| Completed / Fulfilled | Green | `#27AE60` | Text white |
| In Progress | Blue | `#3498DB` | Text white |
| Planned | Light Blue | `#85C1E9` | Text dark |
| Failed | Red | `#E74C3C` | Text white, bold |
| Not Started | Grey | `#D5D8DC` | Text dark |
| Withdrawn | Light Grey | `#BDC3C7` | Strikethrough |

### 11.3 Tab Colours (Summary)

| Worksheet | Tab Colour | Hex Code |
|---|---|---|
| Course Overview | Dark Blue | `#1A5276` |
| Degree Planner | Teal Green | `#117A65` |
| AI Recommendations | Cyan | `#1ABC9C` |
| Weekly Timetable | Purple | `#8E44AD` |
| Academic Calendar | Orange | `#E67E22` |
| Raw Schedule Database | Grey | `#7F8C8D` |

---

## 12. Formula Rules

### 12.1 Formula Catalogue

Each formula is defined conceptually. Implementation uses standard Excel formula syntax.

---

**F-01: Remaining Credits (per requirement)**

| Property | Value |
|---|---|
| Name | `credits_remaining` |
| Purpose | Compute how many credits a degree requirement still needs |
| Definition | `MAX(0, min_credits_required - credits_completed)` |
| Location | `tbl_Degree_Requirements[credits_remaining]` |
| Dependencies | `min_credits_required`, `credits_completed` |

---

**F-02: Credits Completed (per requirement)**

| Property | Value |
|---|---|
| Name | `credits_completed` |
| Purpose | Sum credits earned for courses linked to this requirement that have status `Completed` |
| Definition | `SUMIFS(tbl_Enrolment_Plan[credits_earned], tbl_Enrolment_Plan[requirement_id], [@requirement_id], tbl_Enrolment_Plan[status], "Completed")` |
| Location | `tbl_Degree_Requirements[credits_completed]` |
| Dependencies | `tbl_Enrolment_Plan[credits_earned]`, `tbl_Enrolment_Plan[status]`, `tbl_Enrolment_Plan[requirement_id]` |

---

**F-03: Graduation Progress**

| Property | Value |
|---|---|
| Name | `graduation_progress` |
| Purpose | Overall degree completion as a percentage |
| Definition | `Total_Credits_Completed / Total_Credits_Required` |
| Location | Degree Planner summary row |
| Display format | Percentage, 1 decimal place (e.g., `72.5%`) |
| Dependencies | `SUM(tbl_Degree_Requirements[min_credits_required])`, `SUMIF(tbl_Enrolment_Plan[status], "Completed", tbl_Enrolment_Plan[credits_earned])` |

---

**F-04: Semester GPA Estimate**

| Property | Value |
|---|---|
| Name | `semester_gpa` |
| Purpose | Weighted average grade points for a given semester |
| Definition | `SUMPRODUCT(credits_earned_filtered, grade_points_filtered) / SUM(credits_earned_filtered)` where `filtered` = rows with `semester_code = target` AND `status = "Completed"` |
| Location | Degree Planner summary row |
| Dependencies | `tbl_Enrolment_Plan[credits_earned]`, `tbl_Enrolment_Plan[grade_points]`, `tbl_Enrolment_Plan[semester_code]` |

---

**F-05: Cumulative GPA**

| Property | Value |
|---|---|
| Name | `cumulative_gpa` |
| Purpose | Overall GPA across all completed semesters |
| Definition | Same as Semester GPA but without semester code filter |
| Location | Degree Planner summary row |

---

**F-06: Weekly Class Hours**

| Property | Value |
|---|---|
| Name | `weekly_class_hours` |
| Purpose | Total contact hours per week for a given set of courses |
| Definition | `SUM(tbl_Raw_Schedule[duration_minutes]) / 60` for sessions with `week_pattern="Every Week"` (adjust odd/even/custom proportionally) |
| Location | Degree Planner summary row (current semester) |

---

**F-07: Daily Class Hours**

| Property | Value |
|---|---|
| Name | `daily_class_hours` |
| Purpose | Total hours on a specific day |
| Definition | `SUMIFS(tbl_Raw_Schedule[duration_minutes], tbl_Raw_Schedule[day], "Monday", ...) / 60` |
| Location | Weekly Timetable footer row per day |

---

**F-08: Free Time Hours (per day)**

| Property | Value |
|---|---|
| Name | `free_time_hours` |
| Purpose | Hours with no scheduled sessions between 08:00–18:00 |
| Definition | `10 - daily_class_hours` (for 08:00–18:00 window; adjustable) |
| Location | Weekly Timetable footer row |

---

**F-09: Conflict Detection**

| Property | Value |
|---|---|
| Name | `conflict_flag` |
| Purpose | Flag rows where two sessions overlap in time on the same day in the same semester |
| Definition | `COUNTIFS(other sessions) > 0` where `day = this.day` AND `start_time < this.end_time` AND `end_time > this.start_time` AND `semester_code = this.semester_code` AND `session_id <> this.session_id` |
| Location | Hidden column in `tbl_Raw_Schedule` |

---

**F-10: Time Overlap Detection (Timetable Grid)**

| Property | Value |
|---|---|
| Name | `grid_overlap` |
| Purpose | Highlight timetable grid cells where two or more sessions occupy the same time slot |
| Definition | For each grid cell, `COUNTIFS(tbl_Raw_Schedule[day], grid_day, tbl_Raw_Schedule[start_time], "<="&grid_time, tbl_Raw_Schedule[end_time], ">"&grid_time) > 1` |
| Visual output | Cell turns red with yellow text |

---

**F-11: Duplicate Course Detection**

| Property | Value |
|---|---|
| Name | `duplicate_course_flag` |
| Purpose | Warn if the same course is planned for the same semester twice |
| Definition | `COUNTIFS(tbl_Enrolment_Plan[course_code], [@course_code], tbl_Enrolment_Plan[semester_code], [@semester_code], tbl_Enrolment_Plan[plan_id], "<>"&[@plan_id]) > 0` |
| Location | Conditional formatting on `tbl_Enrolment_Plan` |

---

**F-12: Credit Overflow Warning**

| Property | Value |
|---|---|
| Name | `credit_overflow_warning` |
| Purpose | Warn if planned credits for a semester exceed a reasonable maximum |
| Definition | `SUMIFS(tbl_Enrolment_Plan[credits_earned], tbl_Enrolment_Plan[semester_code], target_semester) > max_credits_per_semester` |
| Default max | 20 (US), 30 (AU), 35 (ECTS), 70 (UK CATS) — adjustable per institution |
| Visual output | Red border on summary cell |

---

**F-13: Prerequisite Not Met Detection**

| Property | Value |
|---|---|
| Name | `prereq_not_met` |
| Purpose | Flag courses where one or more prerequisites have not been completed |
| Definition | For each planned course, extract prerequisites from Course Overview, split by `;`, check each prerequisite code exists in Enrolment Plan with `status="Completed"` |
| Implementation | Helper column or conditional formatting; requires TEXTSPLIT + COUNTIF pattern |

---

**F-14: Course Name Lookup**

| Property | Value |
|---|---|
| Name | `course_name_lookup` |
| Purpose | Display course name alongside course code in other worksheets |
| Definition | `XLOOKUP([@course_code], tbl_Course_Overview[course_code], tbl_Course_Overview[course_name], "Unknown Course")` |
| Location | Helper columns where needed |

---

## 13. Timetable Layout Rules

### 13.1 Grid Dimensions

| Parameter | Default | Adjustable |
|---|---|---|
| Earliest time slot | 08:00 | Yes — institution preference |
| Latest time slot | 21:30 | Yes — institution preference |
| Time increment | 30 minutes | Yes — 15, 30, or 60 min |
| Days included | Monday–Friday | Yes — add Saturday/Sunday |
| Rows in grid | 28 (14 hours × 2 slots) | Computed from time range ÷ increment |

### 13.2 Cell Merging Rules

| Condition | Action |
|---|---|
| Single session occupying N consecutive 30-min slots | Merge N rows in that day's column |
| Two distinct sessions in adjacent slots (different courses) | Do NOT merge; keep separate |
| Back-to-back sessions of the same course but different session types | Do NOT merge (e.g., Lecture then Lab) |
| Sessions crossing the hour boundary (e.g., 09:30–10:30) | Merge 2 slots (09:30 + 10:00) |

### 13.3 Recurring Pattern Support

| Pattern | Logic |
|---|---|
| Every Week | Session appears in every week of the teaching period |
| Odd Weeks | Week number % 2 == 1; session appears in weeks 1, 3, 5, ... |
| Even Weeks | Week number % 2 == 0; session appears in weeks 2, 4, 6, ... |
| Custom Weeks | Session appears only in explicitly listed week numbers |
| Week Range | Session appears in weeks [start_week, end_week] inclusive |
| Single Exception | Excluded dates override all patterns |

### 13.4 Multiple Sessions for One Course

A single course (e.g., MATH101) may generate multiple rows in Raw Schedule Database:
- MATH101 Lecture: Monday 09:00–10:00, Weeks 1–12
- MATH101 Tutorial A: Wednesday 14:00–15:00, Weeks 2–12 (no Week 1)
- MATH101 Lab: Friday 11:00–13:00, Odd Weeks 1–11

All three appear independently in the timetable, each as its own merged block.

### 13.5 Overlapping Detection (Visual)

When two sessions overlap:
1. The overlapping time slot cells are filled red
2. The text `⚠ CONFLICT: {course1} vs {course2}` is displayed
3. A helper note column in Raw Schedule Database logs the conflicting `session_id`

---

## 14. Calendar Export Rules

### 14.1 Supported Export Formats

| Format | Standard | Usage |
|---|---|---|
| Apple Calendar | RFC 5545 `.ics` | Native import; recurring events, alarms, timezone support |
| Google Calendar | RFC 5545 `.ics` | Import via Google Calendar web UI or API |
| Microsoft Outlook | RFC 5545 `.ics` | Import via File → Open → Import; also `.csv` for legacy |
| CSV | Comma-separated | Generic import for any calendar; loses recurrence info |
| JSON | Custom schema | Programmatic consumption; full fidelity |
| PDF | Print layout | Read-only view of timetable |

### 14.2 ICS Export Mapping

| .ics Property | Source Column | Notes |
|---|---|---|
| `UID` | `tbl_Academic_Calendar[calendar_uid]` or generated | Globally unique |
| `DTSTART` | `start_date` + `start_time` | Combined into ISO 8601 datetime |
| `DTEND` | `end_date` + `end_time` | Combined into ISO 8601 datetime |
| `SUMMARY` | `event_name` | Truncated to 255 chars |
| `DESCRIPTION` | `description` | Truncated to 2000 chars |
| `LOCATION` | `location` | Free text |
| `RRULE` | `recurrence_rule` | RFC 5545 format; only if `is_recurring=TRUE` |
| `BEGIN:VALARM` | Generated from `reminder_minutes` | `-PT{minutes}M` trigger |
| `CATEGORIES` | `event_type` | Mapped to calendar categories |

### 14.3 Recurring Event Handling

For Raw Schedule Database sessions exported to calendar:
- Each unique `(course_code, session_type, day, start_time, end_time)` combination generates one recurring `.ics` event
- RRULE: `FREQ=WEEKLY;INTERVAL=1;BYDAY={two-letter day};COUNT={number of weeks in range}`
- Excluded dates: `EXDATE` entries
- Custom weeks: Individual `VEVENT` entries (no RRULE) for each date

### 14.4 Timezone

| Property | Value |
|---|---|
| Default timezone | Institution local (e.g., `Australia/Melbourne`, `America/New_York`, `Asia/Shanghai`) |
| ICS timezone | `VTIMEZONE` block included in `.ics` |
| UTC offset | Automatically derived from timezone |
| Daylight saving | Handled by `VTIMEZONE` definition |

### 14.5 Reminder / Alarm Defaults

| Event Type | Default Reminder |
|---|---|
| Exam | 1 day before + 1 hour before |
| Lecture/Tutorial/Lab | 15 minutes before |
| Assignment deadline | 1 week before + 1 day before |
| Orientation | 1 day before |
| General event | No default reminder |

---

## 15. Print Settings

### 15.1 Global Print Configuration

| Property | Value |
|---|---|
| Paper size | A4 |
| Orientation | Landscape |
| Scaling | Fit to 1 page wide; height automatic |
| Margins (top/bottom/left/right) | 1.5 cm / 1.5 cm / 1.5 cm / 1.5 cm |
| Header | `&[File] — &[Tab]` |
| Footer | `Page &[Page] of &[Pages] | Generated &[Date]` |
| Gridlines | Printed |
| Row/column headings | Not printed |

### 15.2 Per-Worksheet Print Settings

**Course Overview:**
| Property | Value |
|---|---|
| Orientation | Landscape |
| Scaling | Fit all columns to 1 page wide |
| Print area | Table + header |

**Degree Planner:**
| Property | Value |
|---|---|
| Orientation | Landscape |
| Scaling | Fit to 1 page wide |
| Print area | Both tables |
| Page break | After Section A (Degree Requirements) |

**Weekly Timetable:**
| Property | Value |
|---|---|
| Orientation | Landscape |
| Scaling | Fit to 1 page |
| Print area | `A1:H57` |
| Repeat rows at top | Row 1 (day headers) |
| Repeat columns at left | Column A (time labels) |

**Academic Calendar:**
| Property | Value |
|---|---|
| Orientation | Portrait (list format) |
| Scaling | Fit to 1 page wide |
| Print area | Full table |

---

## 16. Data Validation Dictionary

### 16.1 Dropdown Lists (Master)

| List Name | Values | Used In |
|---|---|---|
| `list_course_type` | `Required`, `Major Elective`, `General Education`, `Free Elective` | CO-06 |
| `list_session_type` | `Lecture`, `Tutorial`, `Lab`, `Workshop`, `Seminar`, `Exam` | RS-03 |
| `list_priority` | `Critical`, `High`, `Medium`, `Low` | EP-06, AI-06 |
| `list_difficulty` | `Very Hard`, `Hard`, `Normal`, `Easy` | CO-17 |
| `list_credit_system` | `US Credits`, `ECTS`, `Australian Credit Points`, `UK Credits (CATS)`, `Canadian Credits`, `Singapore Modular Credits`, `Hong Kong Credits`, `Chinese Credits`, `Other` | CO-05 |
| `list_level` | `Introductory`, `Intermediate`, `Advanced`, `Honours`, `Masters`, `Doctoral` | CO-09 |
| `list_status_plan` | `Planned`, `Enrolled`, `In Progress`, `Completed`, `Withdrawn`, `Failed`, `Deferred` | EP-05 |
| `list_decision` | `Pending`, `Accepted`, `Rejected`, `Snoozed` | AI-09 |
| `list_recommendation_type` | `Requirement Fulfilment`, `Interest Match`, `Career Path`, `Prerequisite Chain`, `Workload Balance`, `Popular Course`, `Peer Recommendation`, `Academic Advisor` | AI-04 |
| `list_event_type` | `Semester Start`, `Semester End`, `Teaching Start`, `Teaching End`, `Exam Period`, `Break`, `Holiday`, `Census Date`, `Add/Drop Deadline`, `Withdrawal Deadline`, `Results Release`, `Orientation`, `Graduation`, `Other` | AC-03 |
| `list_week_pattern` | `Every Week`, `Odd Weeks`, `Even Weeks`, `Custom` | RS-12 |
| `list_days` | `Monday`, `Tuesday`, `Wednesday`, `Thursday`, `Friday`, `Saturday`, `Sunday` | RS-04 |
| `list_boolean` | `TRUE`, `FALSE` | CO-23, AC-09 |

### 16.2 Cross-Sheet Validation Rules

| Rule ID | Applies To | Formula |
|---|---|---|
| V-FK-01 | `tbl_Enrolment_Plan[course_code]` | `=COUNTIF(ws1_course_code, [@course_code])>0` |
| V-FK-02 | `tbl_Enrolment_Plan[requirement_id]` | `=COUNTIF(tbl_Degree_Requirements[requirement_id], [@requirement_id])>0` |
| V-FK-03 | `tbl_AI_Recommendations[course_code]` | `=COUNTIF(ws1_course_code, [@course_code])>0` |
| V-FK-04 | `tbl_Raw_Schedule[course_code]` | `=COUNTIF(ws1_course_code, [@course_code])>0` |
| V-CHK-01 | `tbl_Enrolment_Plan` | No duplicate `(course_code, semester_code)` for same student when `status` is not `Withdrawn` or `Failed` |
| V-CHK-02 | `tbl_Raw_Schedule` | `end_time > start_time` |
| V-CHK-03 | `tbl_Raw_Schedule` | `end_week >= start_week` |
| V-CHK-04 | `tbl_Academic_Calendar` | `end_date >= start_date` |
| V-CHK-05 | `tbl_Academic_Calendar` | `end_time > start_time` when both are present |

---

## 17. Conditional Formatting Dictionary (Master)

A unified reference for all conditional formatting rules across the workbook.

| ID | Worksheet | Priority | Condition | Format |
|---|---|---|---|---|
| CF-CO-01 | Course Overview | 1 | `course_type="Required"` | `#8B0000` fill, white text |
| CF-CO-02 | Course Overview | 2 | `course_type="Major Elective"` | `#228B22` fill, white text |
| CF-CO-03 | Course Overview | 3 | `enrolment_difficulty="Very Hard"` | `#FF0000` border, bold |
| CF-CO-04 | Course Overview | 4 | `enrolment_difficulty="Hard"` | `#E67E22` border |
| CF-CO-05 | Course Overview | 5 | `student_rating>=4.5` | `#FFD700` fill |
| CF-CO-06 | Course Overview | 6 | `is_active=FALSE` | Grey text, strikethrough |
| CF-DP-01 | Degree Planner (A) | 1 | `status="Fulfilled"` | `#27AE60` fill, white text |
| CF-DP-02 | Degree Planner (A) | 2 | `status="Not Started"` | `#D5D8DC` fill |
| CF-DP-03 | Degree Planner (A) | 3 | `credits_remaining>0` | `#F1C40F` fill |
| CF-DP-04 | Degree Planner (B) | 4 | `status="Failed"` | `#E74C3C` fill, white, bold |
| CF-DP-05 | Degree Planner (B) | 5 | `status="Completed"` | `#27AE60` fill, white |
| CF-DP-06 | Degree Planner (B) | 6 | `status="Planned"` | `#3498DB` fill, white |
| CF-DP-07 | Degree Planner (B) | 7 | `priority="Critical"` | Dark red text, bold |
| CF-DP-08 | Degree Planner (Summary) | 8 | `graduation_progress<0.5` | `#E74C3C` fill |
| CF-DP-09 | Degree Planner (Summary) | 9 | `graduation_progress>=1` | `#FFD700` fill, bold |
| CF-DP-10 | Degree Planner (Summary) | 10 | `current_semester_credits>max` | `#FF0000` border (3pt) |
| CF-AI-01 | AI Recommendations | 1 | `decision="Accepted"` | `#27AE60` fill, white |
| CF-AI-02 | AI Recommendations | 2 | `decision="Rejected"` | Grey fill, strikethrough |
| CF-AI-03 | AI Recommendations | 3 | `confidence_score>=0.9` | `#FFD700` fill |
| CF-AI-04 | AI Recommendations | 4 | `confidence_score<0.5` | `#FADBD8` fill |
| CF-AI-05 | AI Recommendations | 5 | `expires_at<TODAY()` | Red text, bold |
| CF-TT-01 | Weekly Timetable | 1 | Required course (via VLOOKUP) | `#8B0000` fill, white |
| CF-TT-02 | Weekly Timetable | 2 | Major Elective (via VLOOKUP) | `#228B22` fill, white |
| CF-TT-03 | Weekly Timetable | 3 | General Education (via VLOOKUP) | `#2980B9` fill, white |
| CF-TT-04 | Weekly Timetable | 4 | Cell contains "Lab" | `#8E44AD` fill, white |
| CF-TT-05 | Weekly Timetable | 5 | Cell contains "Tutorial" | `#85C1E9` fill |
| CF-TT-06 | Weekly Timetable | 6 | Cell contains "Workshop" | `#E67E22` fill, white |
| CF-TT-07 | Weekly Timetable | 7 | Cell contains "Exam" | `#FF0000` border (3pt), bold |
| CF-TT-08 | Weekly Timetable | 8 | Cell is blank | `#EAECEE` fill |
| CF-TT-09 | Weekly Timetable | 9 | Overlap detected | `#FF0000` fill, yellow text, bold |
| CF-AC-01 | Academic Calendar | 1 | `event_type="Holiday"` | `#F1C40F` fill |
| CF-AC-02 | Academic Calendar | 2 | `event_type="Exam Period"` | `#E74C3C` fill, white |
| CF-AC-03 | Academic Calendar | 3 | `event_type="Break"` | `#D6EAF8` fill |
| CF-AC-04 | Academic Calendar | 4 | `event_type` contains "Deadline" | `#E67E22` border, bold |
| CF-AC-05 | Academic Calendar | 5 | `end_date<TODAY()` | Grey text |
| CF-AC-06 | Academic Calendar | 6 | `start_date=TODAY()` | Green fill, bold |
| CF-RS-01 | Raw Schedule | 1 | `conflict_flag=TRUE` | `#FF0000` fill, yellow text, bold |
| CF-RS-02 | Raw Schedule | 2 | `session_type="Lecture"` | `#8B0000` fill, white |
| CF-RS-03 | Raw Schedule | 3 | `session_type="Lab"` | `#8E44AD` fill, white |
| CF-RS-04 | Raw Schedule | 4 | `session_type="Tutorial"` | `#2980B9` fill, white |
| CF-RS-05 | Raw Schedule | 5 | `session_type="Workshop"` | `#E67E22` fill, white |
| CF-RS-06 | Raw Schedule | 6 | `session_type="Seminar"` | `#1ABC9C` fill, white |
| CF-RS-07 | Raw Schedule | 7 | `session_type="Exam"` | `#FF0000` border (3pt), bold |
| CF-RS-08 | Raw Schedule | 8 | `week_pattern="Odd Weeks"` | `#FEF9E7` fill |
| CF-RS-09 | Raw Schedule | 9 | `week_pattern="Even Weeks"` | `#E8F8F5` fill |

---

## 18. Versioning & Migration Strategy

### 18.1 Schema Version

| Property | Value |
|---|---|
| Current version | `2.0.0` |
| Schema version stored in | Named range `schema_version` — value in a hidden cell on Course Overview sheet at `Z1` |
| Format | Semantic versioning (`MAJOR.MINOR.PATCH`) |
| Upgrade path | New columns appended to right; deprecated columns hidden, not deleted; migration formula sheet tracks transformations |

### 18.2 Backward Compatibility

- New columns always appended to the right of existing columns in each table.
- Column removal: hide via `Visible=N`, never delete. Mark as deprecated in a migration log.
- Data type changes: add a new column with the new type; populate via formula from old column; hide old column.
- Enum value additions: append to dropdown list; no removal of existing values unless with a major version bump.

### 18.3 Institution Customisation Points

| Customisation | Method |
|---|---|
| Credit system defaults | Cell `credit_system_default` named range; validated against CO-05 enum |
| Max credits per semester | Cell `max_credits_per_semester` named range; adjusts F-12 threshold |
| Timetable start/end time | Cells `timetable_start_hour`, `timetable_end_hour` named ranges |
| Time increment | Cell `timetable_increment_minutes` named range (15, 30, or 60) |
| Days shown | Checkboxes or list: Monday–Sunday in a control area |
| GPA scale | Named range `gpa_scale` with lookup table mapping letter grades to points |
| Institution branding | Header/footer text; logo image in first worksheet header area |
| Local calendar | Default timezone named range; public holiday list appended to Academic Calendar |

### 18.4 AI Generation Compatibility

The schema is designed so that an AI system can:
1. Generate the complete workbook from this specification alone (all structure, validation, conditional formatting, and named ranges are fully specified).
2. Populate Course Overview from a university handbook API or structured data extract.
3. Generate a Degree Planner from degree requirement XML/JSON.
4. Generate a Raw Schedule Database from timetable data feeds.
5. Render the Weekly Timetable from Raw Schedule Database automatically (formulas + conditional formatting only; no VBA).
6. Produce AI Recommendations by comparing Degree Planner gaps against Course Overview.
7. Generate `.ics` calendar exports by iterating Raw Schedule Database and Academic Calendar tables.

### 18.5 Automation Compatibility

- All tables are Excel structured references (`tbl_Name[column]`) — formulas auto-expand.
- No VBA or macros — fully compatible with Google Sheets, Excel Online, LibreOffice.
- All logic expressed via formulas, data validation, and conditional formatting.
- Named ranges provide stable references for programmatic access.
- The schema supports import/export via Power Query, Python (openpyxl / pandas), and Google Apps Script.

---

**End of Specification**

---

*This schema specification defines the complete structure for the AI University Course Planner Excel workbook. An AI system receiving this document as input can generate a functionally identical workbook without ambiguity.*
