# oPenEfficiency Agent Skills Index

This directory (`.agents`) is the **Single Source of Truth (SSOT)** for all AI agents (Claude, Gemini, Copilot, Cursor) working on the oPenEfficiency PowerPoint Add-in.

---

## 1. Mandatory Pre-Task Workflow

**Before starting ANY task**, agents must:

1. **Read `.agents/project-overview.md`** - Understand the architecture and patterns.
2. **Read `.agents/core-guidelines.md`** - Review the strict project mandates (VSTO, UI, Code Style).
3. **Find your Task Type below** - Read the associated skill file.
4. **Read relevant existing code files** - Search for existing implementations before starting.
5. **Then implement.**

**Do NOT skip to implementation without reading relevant documentation.**

---

## 2. Task Type → Skill Mapping

Find your task type below and read the indicated files:

| Task Type | Read File |
|-----------|-----------|
| **Add a new feature** (e.g., button, chart, alignment tool) | `skills/adding-features.md` |
| **Create a dialog/window** (e.g., settings, wizard, popup) | `skills/floating-window-creation.md` |
| **Modify existing UI** (e.g., sidebar, XAML, WPF) | `skills/ui-development.md` |
| **Debug / Fix an error** (e.g., compilation, CS2001, missing connectionId) | `skills/vsto-pitfalls.md` |
| **Cleanup workspace** (e.g., temp files, structure) | `skills/workspace-hygiene.md` |

---

## 3. Available Skills (Deep Dives)

### `adding-features`
**File:** `skills/adding-features.md`
**Triggers:** "add feature", "new feature", "create feature", "add button", "sidebar feature", "toggle feature"
**Purpose:** Complete step-by-step guide for adding new features using the `[FeatureMetadata]` attribute pattern.

### `ui-development`
**File:** `skills/ui-development.md`
**Triggers:** "UI", "dialog", "WPF", "XAML", "modify sidebar", "styling", "theme", "color"
**Purpose:** Guidelines for making consistent UI changes, styling conventions, and DPI-aware positioning.

### `floating-window-creation`
**File:** `skills/floating-window-creation.md`
**Triggers:** "floating window", "floating toolbar", "context menu", "modeless window", "create popup"
**Purpose:** Streamlined skill for creating modern, dark-themed floating WPF windows with drag-to-move.

### `vsto-pitfalls`
**File:** `skills/vsto-pitfalls.md`
**Triggers:** "debug", "error", "not working", "VSTO", "csproj", "compilation", "crash", "InitializeComponent"
**Purpose:** Essential VSTO development rules, identifying `.csproj` globbing issues, and bug prevention.

### `workspace-hygiene`
**File:** `skills/workspace-hygiene.md`
**Triggers:** "cleanup", "temp files", "clean workspace", "trash"
**Purpose:** Maintaining clean repo root, temporary file cleanup (`_temp_scripts/`), folder structure.

---

## 4. Project Documentation Overview

| Document | Purpose |
|----------|---------|
| [`project-overview.md`](./project-overview.md) | Architecture, core patterns, PowerPoint COM handling |
| [`core-guidelines.md`](./core-guidelines.md) | Project mandates, code style, UI/UX rules, stability |
| [`INDEX.md`](./INDEX.md) | (This file) Navigation and skill discovery |