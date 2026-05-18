# GEMINI.md

This file provides guidance to Gemini CLI when working with code in this repository. Instructions here are foundational mandates and take absolute precedence.

## Context Compression & Information Preservation

**Problem:** Long conversations get compressed by the system. Context about decisions, intermediate results, or work-in-progress can be lost.

**Solutions (use these strategies):**

### 1. Memory for Cross-Session Context
Save critical information to the global memory or project-specific storage:
- **User preferences** - How you want to collaborate
- **Project state** - Current goals, deadlines, stakeholder decisions
- **Feedback** - What worked/didn't work, corrections
- **References** - Links to external resources (Linear, Grafana, etc.)

**When to save:** When you learn something that should persist beyond this conversation (decisions, preferences, incidents, goals).

### 2. Checkpoint Files for In-Session State
For work that spans multiple steps within a session, write state to a file:
```csharp
// Example: FeaturesMigrationCheckpoint.md
// Batch 3 complete: 14/16 features migrated
// Remaining: StyleCheckFeature, SyncObjectsFeature (manual dialogs)
// Next: Start Batch 4 (Utilities)
```
**When to use:** Multi-step refactors, batch migrations, or any task where you need to resume mid-work after a context shift.

### 3. Explicit Summarization Before Compression
When you sense context is getting long, ask:
> "Summarize what we've accomplished and what's next before context compresses"

This forces a checkpoint that can be referenced later.

### 4. File-Based State Over Chat History
- **Don't:** Describe complex state in chat messages (gets compressed)
- **Do:** Write state to files (CLAUDE.md, GEMINI.md, checkpoint files, feature lists)

**Example pattern:**
1. Read current state from file
2. Make changes
3. Update file with new state
4. Reference file (not chat history) in future turns

### 5. Task Lists for Progress Tracking
Use the task system to track what's done vs. pending. Tasks persist and provide a "table of contents" for complex work.

### 6. Knowledge Persistence & Bug Documentation
**Mandate:** Whenever a complex bug, compilation error, or architectural mismatch is resolved—especially those specific to the VSTO/WPF integration—the agent **MUST** document the root cause and the permanent fix. 
- Update `.agents/skills/vsto-pitfalls.md` with the lesson learned.
- Ensure the fix is described technically (why it happened) and procedurally (how to prevent it).
- **Never** let a hard-won lesson be lost to context compression.

---

## Additional:

Always ask Questions if the Input from the user is unclear or ambigous. Always answer with "Okay chef".

## Communication & Uncertainty
**Mandate:** If you are unsure about any implementation detail, architectural decision, or the exact scope of a requested change, you **MUST** ask the user for clarification (Rückfragen) before proceeding with file modifications. It is better to confirm intent than to guess and introduce technical debt.

## Versioning Policy

**Automated Naming:** The application version is automatically updated during the GitHub Actions `publish` workflow. The format is **`yyyy.mm.dd.RunNumber`** (Year.Month.Day.RunNumber).
- **No Manual Action Required:** You do NOT need to manually update `<ApplicationVersion>` in `oPenEfficiency.csproj` or `AssemblyVersion` in `Properties\AssemblyInfo.cs`.
- **Auto-Increment:** Ensure `<AutoIncrementApplicationRevision>` is always set to `false` in the `.csproj` to prevent Visual Studio from overriding this scheme locally.

---

## Temporary Files & Scripts

**Mandate:** The project root MUST remain clean at all times. AI agents MUST store all custom scripts (Python, PowerShell, Bash), temporary logs, and intermediate outputs EXCLUSIVELY in the **`_temp_scripts/`** directory.
- **Root Cleanup:** Do not create new scripts, `.txt`, or `.ps1` files directly in the root directory. 
- **Naming:** Use descriptive names for scripts (e.g., `_temp_scripts/update_sidebar_layout.py`) to maintain traceability.
- **Cleanup:** Delete temporary scripts once the task is fully verified, unless they provide lasting value for future automation.

---

## UI/UX & Theming Guidelines

**Theme Cohesion (Light vs. Dark):**
- **Contrast on small elements:** Never use `TextMutedBrush` or `TextSecondaryBrush` for standalone icon buttons or close ("X") buttons in floating windows. Due to optical blending against light surfaces (`SurfaceBrush` in Light Theme), they suffer from severe contrast issues. Always use `TextPrimaryBrush` for such interactive elements.
- **Tooltip Inheritance:** WPF ToolTips natively inherit `Foreground` properties from their parent controls. If a button turns its text white on hover, the ToolTip will inherit white text. Always ensure `ToolTip` styles explicitly set their own `Background` and `Foreground` decoupled from the button hover states.

## Agenda Engine Compatibility Strategy (EE4P & Think-Cell)

**Goal:** Ensure the native oPenEfficiency Agenda Wizard can safely coexist with or integrate with third-party formats.

### Efficient Elements (EE4P) Integration
- **Adapter Pattern:** The `AgendaGenerationService` must be able to serialize our lightweight `AgendaItemData` and native PowerPoint Sections into the specific XML schema expected by EE4P.
- **Tag Injection:** When generating shapes, inject the corresponding `EE4P_AGENDAWIZARD` string tags (e.g., `Topic`, `TimeSlot`) and use their GUID-based naming patterns (e.g., `ee4p_item_[guid]_shape`).

### ThinkCellGuard
- **Safe Coexistence (No Overwrite):** Full compatibility is not feasible due to Think-Cell's proprietary hashing and internal rendering engine.
- **ThinkCellGuard:** The `AgendaService` must actively scan the presentation for `THINKCELLPRESENTATIONDONOTDELETE` and `THINKCELLSHAPEDONOTDELETE` tags. Any slide containing these tags must be skipped and locked from automated oPenEfficiency updates to prevent corruption of the user's Think-Cell links.

---

## Technical Integrity & VSTO Stability

**Critical Mandates for Build & Runtime Stability:**

1. **NO GLOBBING in `.csproj`:** NEVER use wildcards like `**\*.cs` or `**\*.xaml` for UI, Models, Services, or Utils folders. Every file MUST be explicitly listed with its correct `DependentUpon` mapping. Wildcards break VSTO code-behind linking (causing `InitializeComponent` missing errors). The ONLY exception permitted is the `Features` folder.
2. **NO DESTRUCTIVE REFACTORING ON BUILD ERRORS:** If you encounter massive build errors (like `CS0103` or missing `InitializeComponent`), DO NOT assume the C# logic is fundamentally broken and delete/simplify large portions of code. First, check if `.csproj` globbing or XAML structure was recently altered. Isolate `.csproj` changes from C# logic changes to prevent misdiagnosing the root cause.
3. **XAML connectionId Safety:** Define all complex `ContextMenu` and `Resources` as **StaticResources** at the top of the XAML file. NEVER nest anonymous context menus inside ListBox templates as it causes identity collisions in shared namespaces.
3. **Constructor-First Initialization:** Perform all primary UI population and child control instantiation in the **Constructor**, not the `Loaded` event. The `Loaded` event is suppressed in VSTO's `ElementHost` and will cause blank windows.
4. **Scrolling Constraints:** Gallery views (ListBox, WrapPanel) MUST have `HorizontalScrollBarVisibility="Disabled"` to force grid wrapping in the PowerPoint TaskPane.
5. **Office Namespace Precision:** Always use fully qualified names for Office COM types (e.g., `Microsoft.Office.Core.SmartArtNode`) to avoid name collisions with standard .NET types.
6. **Strict XAML Setter Syntax:** ALWAYS use the explicit `<Setter Property="X" Value="Y" />` syntax. NEVER use shorthand attributes like `<Setter X="Y" />` in Styles or Templates, as the VSTO build task will fail with `MC3072`.

