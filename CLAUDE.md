# CLAUDE.md

This file serves as a lightweight dispatcher for Claude Code (`claude.ai/code`) when working in this repository.

**For all architecture, design, UI, and project-specific guidelines, refer to the Single Source of Truth (SSOT) in the `.agents` folder.**

---

## Mandatory Pre-Task Workflow

Before starting any task, you **MUST** consult the AI Agent index:

1. **Read `.agents/INDEX.md`** to find the correct guidelines and skills for your task.
2. **Read `.agents/core-guidelines.md`** for strict project mandates (VSTO stability, UI rules, code style).
3. **Read `.agents/project-overview.md`** for architectural patterns.

---

## Build & Run

```bash
# Open solution in Visual Studio 2022
oPenEfficiency/oPenEfficiency.slnx

# Build (requires Visual Studio with Office Developer Tools)
msbuild oPenEfficiency/oPenEfficiency.slnx /p:Configuration=Release /p:Platform="Any CPU"

# Or use Visual Studio: Ctrl+Shift+B
```

**Prerequisites:**
- Visual Studio 2022 with Office development workload
- .NET Framework 4.8 SDK
- PowerPoint 2016+ installed (for debugging)

**Debugging:** Press F5 in Visual Studio - PowerPoint launches with the add-in loaded.