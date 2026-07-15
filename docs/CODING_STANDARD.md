# ACC KPI Framework

# Coding Standard

Version: 2.0.0-alpha

Status: Approved

Last Updated: 2026-07-15

---

# 1. Purpose

This document defines mandatory development rules for the ACC KPI Framework.

These rules apply to:

- ChatGPT
- Codex
- GitHub Copilot
- Human developers

Compliance with this standard is mandatory.

---

# 2. General Principles

## CS-001

Architecture has higher priority than implementation.

---

## CS-002

Code quality has higher priority than implementation speed.

---

## CS-003

Any architectural decision must be documented.

---

## CS-004

Every change must be traceable.

---

## CS-005

One module = one responsibility.

---

# 3. Project Architecture

Project layers:

1. Constants
2. Layout
3. Geometry
4. Builder
5. Calculations
6. Navigation
7. Events
8. Business Logic

Each layer has a single responsibility.

---

# 4. Constants

Only immutable values belong in CONST.

Examples:

- sheet names
- statuses
- result types
- formats
- colors
- default values

Forbidden:

- geometry
- calculations
- SpreadsheetApp

---

# 5. Layout

LAYOUT is the single source of truth for document geometry.

LAYOUT contains:

- row indexes
- column indexes
- merged ranges
- block sizes
- row heights
- headers
- widths
- offsets

Forbidden:

- calculations
- SpreadsheetApp
- business logic

---

# 6. Geometry

Geometry converts LAYOUT into usable coordinates.

Allowed:

- coordinate calculations
- block selection
- helper methods

Forbidden:

- business calculations
- UI creation

---

# 7. Builder

Builder creates document structure.

Allowed:

- formatting
- borders
- merged cells
- drawing

Forbidden:

- calculations
- business logic

---

# 8. Calculations

Calculation modules perform calculations only.

Forbidden:

- hardcoded coordinates
- formatting
- merged cells
- UI creation

Calculations use Geometry API.

---

# 9. Navigation

Navigation contains only navigation logic.

No calculations.

No formatting.

---

# 10. Events

Event modules contain only event handlers.

Business logic must be delegated.

---

# 11. Magic Numbers

Magic numbers are forbidden.

Forbidden:

```javascript
sheet.getRange(78,27)
```

Correct:

```javascript
const sprint = getSprintLayoutV1_(month, sprint);

sheet.getRange(
    sprint.firstHypothesisRow,
    LAYOUT.HYPOTHESIS.KPI_COLUMN
);
```

---

# 12. String Literals

Business literals are forbidden.

Forbidden:

```javascript
"Completed"

"Success"

"Quarter"

"Sprint"

"ACC_TEMPLATE"
```

Correct:

```javascript
CONST.STATUS.COMPLETED

CONST.RESULT.SUCCESS

CONST.SHEETS.TEMPLATE
```

---

# 13. Public API

Public functions:

get...

build...

draw...

calculate...

update...

validate...

clear...

reset...

sync...

Private helpers end with "_".

Example:

calculateMonthPlan()

calculateMonthPlan_()

---

# 14. Function Size

Recommended:

up to 80 lines

Maximum:

150 lines

Longer functions must be split.

---

# 15. Documentation

Every public function must contain JSDoc.

Example:

/**
 * Calculates sprint totals.
 *
 * @param {Sheet} sheet
 * @returns {void}
 */

---

# 16. Code Style

Use:

const

instead of

var

Prefer:

const

over

let

when possible.

Indentation:

2 spaces.

Maximum line length:

100 characters.

---

# 17. Module Structure

Each module contains:

1. Header
2. Public functions
3. Private helpers

Order is mandatory.

---

# 18. Dependencies

Circular dependencies are forbidden.

Builder must not depend on Calculations.

Calculations must not depend on Builder.

---

# 19. Git Workflow

One task = one branch.

One logical change = one commit.

Commit format:

ARCH-001 Complete Layout

CALC-003 New KPI algorithm

FIX-004 Progress calculation

---

# 20. Development Workflow

Task

↓

Architecture Review

↓

Implementation

↓

Code Review

↓

Testing

↓

Commit

↓

clasp push

↓

Google Sheets Test

↓

Merge

Skipping stages is forbidden.

---

# 21. Testing

Every completed task must include testing.

Minimum:

□ Existing functionality

□ New functionality

□ Regression

---

# 22. Documentation

Architecture changes require updates to:

CHANGELOG

TASKS

DECISIONS

Architecture

---

# 23. AI Development Rules

AI must provide complete artifacts.

Forbidden:

- examples instead of final files
- partial implementations
- incomplete modules

Required:

- complete files
- complete functions
- complete configurations
- complete documentation

---

# 24. AI Workflow

Every AI response must include:

1. Analysis
2. Affected files
3. Risks
4. Implementation
5. Tests
6. Commit message
7. Documentation changes

---

# 25. Architecture Review Checklist

Before merge verify:

□ No magic numbers

□ No business literals

□ No duplicated code

□ Layout used correctly

□ Geometry used correctly

□ Builder contains no calculations

□ Calculations contain no geometry

□ Documentation updated

---

# 26. Definition of Done

Task is complete only if:

□ Code implemented

□ Tests passed

□ Documentation updated

□ Git committed

□ Apps Script updated

□ Google Sheets tested

---

# 27. AI Interaction Standard

AI acts as project architect.

Responsibilities:

- architecture
- design
- review
- technical debt control
- documentation
- planning

AI does not produce temporary solutions unless explicitly requested.

AI always prefers maintainable architecture.

---

# 28. Project Philosophy

ACC KPI Framework is a software framework.

It is not a collection of Google Apps Script files.

Every decision must improve:

- maintainability
- readability
- scalability
- testability
- extensibility

Long-term quality has higher priority than short-term implementation speed.

---

END OF DOCUMENT
