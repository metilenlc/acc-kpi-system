# ACC KPI Framework

## Project Goal

Framework for KPI management built on Google Sheets + Apps Script.

---

## Architecture

Single source of truth:

- CONST
- LAYOUT
- Geometry API

---

## Coding Rules

- No magic numbers.
- No hardcoded sheet coordinates.
- No business string literals.
- One module = one responsibility.
- Geometry only through LAYOUT.
- Statuses only through CONST.

---

## Development Workflow

Architecture
↓

Implementation

↓

Test

↓

Commit

↓

clasp push

↓

Google Sheets test

---

## Git

One task = one branch.

One logical change = one commit.
