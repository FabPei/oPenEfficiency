# GEMINI.md

This file serves as a lightweight dispatcher for Gemini CLI when working in this repository.

**For all architecture, design, UI, and project-specific guidelines, refer to the Single Source of Truth (SSOT) in the `.agents` folder.**

---

## Mandatory Pre-Task Workflow

Before starting any task, you **MUST** consult the AI Agent index:

1. **Read `.agents/INDEX.md`** to find the correct guidelines and skills for your task.
2. **Read `.agents/core-guidelines.md`** for strict project mandates (VSTO stability, UI rules, code style).
3. **Read `.agents/project-overview.md`** for architectural patterns.

---

## Context Compression & Memory Rules (Gemini Specific)

**Problem:** Long conversations get compressed by the system. Context about decisions, intermediate results, or work-in-progress can be lost.

### Strategies to Use:
1. **Memory for Cross-Session Context:** Save critical information (preferences, project state, decisions) to global memory or project-specific storage.
2. **Checkpoint Files:** For multi-step tasks, write progress state to a file (e.g., `_temp_scripts/checkpoint.md`) instead of relying on chat history.
3. **Explicit Summarization:** Ask the user to summarize before context compresses.
4. **Document Bugs:** If you resolve a complex VSTO/WPF bug, you **MUST** document the root cause and the permanent fix in `.agents/skills/vsto-pitfalls.md`. Never let a hard-won lesson be lost.

Always ask questions if the input from the user is unclear or ambiguous. Always answer with "Okay chef".