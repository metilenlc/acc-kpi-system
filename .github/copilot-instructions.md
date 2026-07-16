# ACC KPI Framework

> AI Development Instructions

## Project

ACC KPI Framework is a Google Apps Script framework for KPI planning,
analytics and sprint management inside Google Sheets.

This repository is treated as a software framework rather than a collection
of scripts.

---

# Architecture First

Architecture always has higher priority than implementation.

Never modify code before understanding:

- purpose
- dependencies
- public API
- affected modules

Every architectural change must be reviewed before implementation.

---

# Single Source of Truth

Never duplicate project structure.

Use only:

- CONST
- LAYOUT
- Geometry API

Never hardcode:

- row indexes
- column indexes
- sheet names
- block sizes
- business statuses
- business literals

---

# Module Responsibilities

## Constants

Contains only immutable constants.

Never store geometry.

---

## Layout

Contains the complete document geometry.

Never contains calculations.

Never accesses SpreadsheetApp.

---

## Geometry

Provides API for accessing the layout.

Responsible for translating LAYOUT into usable coordinates.

---

## Builder

Creates document structure.

Never performs calculations.

Never contains business logic.

---

## Calculations

Contains only calculations.

Never knows sheet geometry.

Never creates UI.

---

## Navigation

Contains navigation only.

---

## Events

Contains only event handlers.

---

# Coding Rules

No magic numbers.

No duplicated code.

No hidden dependencies.

No global mutable state.

Keep functions small.

Prefer descriptive names.

Public functions must be documented.

---

# Naming Convention

Functions:

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

---

# Git Workflow

One task = one branch.

One logical change = one commit.

Commit message format:

ARCH-001 Complete Layout refactoring

CALC-003 New KPI distribution

FIX-005 Sprint progress bug

---

# Development Workflow

1. Analyze
2. Design
3. Implement
4. Review
5. Test
6. Commit
7. Push
8. Update documentation

Never skip review.

---

# Testing

Every completed task must be tested inside Google Sheets.

Never assume functionality works.

---

# Documentation

Keep documentation synchronized with implementation.

Update:

- CHANGELOG
- TASKS
- DECISIONS

when architecture changes.

---

# Current Architecture

Version:

2.0.0-alpha

Current Sprint:

Architecture v2.0

Current Epic:

ARCH-001
Architecture Refactoring

---

# Project Goal

Create a reusable Google Apps Script framework for KPI planning,
sprint management and analytics.

The framework must be scalable,
maintainable
and easy to extend.

Code quality has higher priority than implementation speed.
