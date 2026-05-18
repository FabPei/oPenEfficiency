---
name: workspace-hygiene
description: "cleanup, clean workspace, temp files, delete scripts, unorganized, structure, folder cleanup, trash"
---

# Workspace Hygiene & Repository Cleanliness

Maintaining a clean project repository is critical to prevent context bloat, reduce token consumption, and ensure that the root directory remains easy to navigate. AI agents MUST follow these rigorous hygiene rules during every session.

---

## 1. Zero-Tolerance for Root Clutter

**Rule:** The project root directory MUST NOT be used for temporary scripts, dumps, logs, or unorganized text files.
- **NEVER** create `.txt`, `.py`, `.ps1`, or `.md` files directly in the root directory unless they are officially part of the project documentation (like `README.md` or `CHANGELOG.md`).
- **Do not** leave screenshots or arbitrary output logs (e.g., `build_output.txt`) in the root.

## 2. Managing the `_temp_scripts/` Graveyard

**Rule:** While `_temp_scripts/` is the designated location for intermediate scripts, it must not become a dumping ground.
- **Proactive Cleanup:** At the end of a successful task, feature implementation, or bugfix, the agent MUST explicitly delete any "throwaway" scripts (e.g., `apply_fix.py`, `test_x.ps1`) created during that session.
- **Exceptions:** Scripts that provide lasting value for recurring automation tasks (e.g., a complex refactoring script that might be needed again) can be kept, but they should be given highly descriptive names.

## 3. Ephemeral Checkpoints & Plans

**Rule:** Checkpoints and plans are temporary aids for the agent's context and must be cleaned up when the task is done.
- If a plan is created (e.g., `ThemeRefactoringPlan.md`), it should be deleted or moved to a historical documentation folder once the refactoring is 100% complete and verified.
- Avoid leaving behind numbered or phased files (like `Phase1_Research_Translation.md`) in the root.

## 4. Unstructured Input Data

**Rule:** Random specifications, duplicate text files, or raw input data MUST NOT be stored loosely in the root.
- If raw documentation or notes are provided by the user (e.g., `Anforderungsanalyse EmpowerSuite Funktionen.txt`), store them in an appropriate documentation directory like `various input/` or a dedicated `docs/` folder.
- Deduplicate files actively. If a file exists in `various input/`, do not keep a copy in the root.

---

## Agent Enforcement Workflow

Whenever an agent finishes a multi-turn task, they MUST:
1. Verify if they created any temporary files or scripts to execute the task.
2. Delete them immediately via `run_shell_command` using `Remove-Item -Force`.
3. Inform the user that the temporary workspace artifacts have been cleaned up.