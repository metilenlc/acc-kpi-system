# ACC KPI Framework

Version: 2.0.0-alpha

---

## Description

ACC KPI Framework is a modular Google Apps Script framework for KPI planning,
analytics, sprint management and account management inside Google Sheets.

The framework is designed for long-term development and follows software engineering practices including version control, architectural documentation and modular design.

---

# Technology Stack

- Google Sheets
- Google Apps Script
- clasp
- Visual Studio Code
- Git
- GitHub
- OpenAI Codex
- ChatGPT

---

# Repository Structure

```
ACC_KPI_SYSTEM/
│
├── .github/
│   └── copilot-instructions.md
│
├── .vscode/
│
├── docs/
│   ├── README.md
│   ├── ARCHITECTURE.md
│   ├── CODING_STANDARD.md
│   ├── CHANGELOG.md
│   ├── DECISIONS.md
│   ├── PROJECT_STATE.md
│   ├── RELEASES.md
│   ├── ROADMAP.md
│   └── TASKS.md
│
├── 00_Constants.js
├── 01_LayoutMap.js
├── 02_Ranges.js
├── 03_Geometry.js
├── 04_Build.js
├── 05_Dropdowns.js
├── 06_Calendar.js
├── 07_Calculations.js
├── 08_Progress.js
├── 09_Menu.js
├── 10_Events.js
├── 11_Accounts.js
├── 12_Utils.js
├── 13_Groups.js
├── 14_Validation.js
├── 15_Directory.js
├── 16_BuilderParts.js
│
├── .clasp.json
├── appsscript.json
├── .gitignore
└── README.md
```

---

# Project Principles

- Architecture first.
- No magic numbers.
- No duplicated business literals.
- Geometry is defined only in Layout.
- Builder contains no business logic.
- Calculations contain no template geometry.
- One module = one responsibility.

---

# Development Workflow

1. Select task from `docs/TASKS.md`
2. Create Git branch
3. Architecture analysis
4. Implementation
5. Testing
6. Commit
7. `clasp push`
8. Google Sheets validation
9. Merge into `main`

---

# Branch Naming

```
feature/<name>

refactor/<name>

fix/<name>

docs/<name>
```

Examples:

```
feature/month-summary

refactor/layout-v2

fix/progress

docs/readme
```

---

# Commit Message Format

```
ARCH-001 Complete Geometry API

CALC-002 Monthly KPI distribution

BUILD-003 Dynamic sprint rows

FIX-004 Progress calculation

DOC-001 Add project README
```

---

# Documentation

All technical documentation is located in the `docs` directory.

| Document           | Purpose                |
| ------------------ | ---------------------- |
| README.md          | Documentation overview |
| ARCHITECTURE.md    | System architecture    |
| CODING_STANDARD.md | Development standards  |
| TASKS.md           | Current sprint         |
| PROJECT_STATE.md   | Current project status |
| ROADMAP.md         | Planned functionality  |
| DECISIONS.md       | Architecture decisions |
| CHANGELOG.md       | Change history         |
| RELEASES.md        | Release history        |

The `docs` directory is the single source of truth for project documentation.

---

# Current Status

Version:

2.0.0-alpha

Current Sprint:

Architecture v2.0

Current Epic:

ARCH-001

---

# License

Internal project.

All rights reserved.
