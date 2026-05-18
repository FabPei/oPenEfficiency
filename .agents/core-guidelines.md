# Core Guidelines & Mandates

This document serves as the **Single Source of Truth (SSOT)** for all critical architecture, design, and coding rules in the oPenEfficiency project. **These instructions take absolute precedence over any generic AI behaviors.**

---

## 1. Technical Integrity & VSTO Stability

**Critical Mandates for Build & Runtime Stability:**

1. **NO GLOBBING in `.csproj`:** NEVER use wildcards like `**\*.cs` or `**\*.xaml` for UI, Models, Services, or Utils folders. Every file MUST be explicitly listed with its correct `DependentUpon` mapping. Wildcards break VSTO code-behind linking (causing `InitializeComponent` missing errors). The ONLY exception permitted is the `Features` folder.
2. **NO DESTRUCTIVE REFACTORING ON BUILD ERRORS:** If you encounter massive build errors (like `CS0103` or missing `InitializeComponent`), DO NOT assume the C# logic is fundamentally broken and delete/simplify large portions of code. First, check if `.csproj` globbing or XAML structure was recently altered. Isolate `.csproj` changes from C# logic changes to prevent misdiagnosing the root cause.
3. **XAML connectionId Safety:** Define all complex `ContextMenu` and `Resources` as **StaticResources** at the top of the XAML file. NEVER nest anonymous context menus inside ListBox templates as it causes identity collisions in shared namespaces.
4. **Constructor-First Initialization:** Perform all primary UI population and child control instantiation in the **Constructor**, not the `Loaded` event. The `Loaded` event is suppressed in VSTO's `ElementHost` and will cause blank windows.
5. **Scrolling Constraints:** Gallery views (ListBox, WrapPanel) MUST have `HorizontalScrollBarVisibility="Disabled"` to force grid wrapping in the PowerPoint TaskPane.
6. **Office Namespace Precision:** Always use fully qualified names for Office COM types (e.g., `Microsoft.Office.Core.SmartArtNode`) to avoid name collisions with standard .NET types.
7. **Strict XAML Setter Syntax:** ALWAYS use the explicit `<Setter Property="X" Value="Y" />` syntax. NEVER use shorthand attributes like `<Setter X="Y" />` in Styles or Templates, as the VSTO build task will fail with `MC3072`.

---

## 2. Code Style & Implementation Rules

### Error Handling
```csharp
catch (Exception ex)
{
    ExceptionLogger.Log(ex, "ClassName.MethodName");
    // Don't swallow exceptions silently
}
```

### PowerPoint Access
```csharp
var manager = new PowerPointManager(Globals.ThisAddIn.Application);
// Never use: var app = Globals.ThisAddIn.Application; app.ActiveWindow...
```

### COM Objects
```csharp
// 1-indexed collections in COM!
var shape = shapes[1]; // NOT shapes[0]
```

### Shape Metadata
```csharp
// Use oPE_ prefix for custom tags
shape.Tags.Add("oPE_Widget", "YourFeature");

// Use AlternativeText for state (survives copy/paste)
group.AlternativeText = $"{FeatureTag}|{value}";
```

---

## 3. UI/UX & Theming Guidelines

**Theme Cohesion (Light vs. Dark):**
- **Contrast on small elements:** Never use `TextMutedBrush` or `TextSecondaryBrush` for standalone icon buttons or close ("X") buttons in floating windows. Due to optical blending against light surfaces (`SurfaceBrush` in Light Theme), they suffer from severe contrast issues. Always use `TextPrimaryBrush` for such interactive elements.
- **Tooltip Inheritance:** WPF ToolTips natively inherit `Foreground` properties from their parent controls. If a button turns its text white on hover, the ToolTip will inherit white text. Always ensure `ToolTip` styles explicitly set their own `Background` and `Foreground` decoupled from the button hover states.
- **WPF Colors:** NEVER hardcode colors (e.g., `Background="#252525"`). ALWAYS use `{StaticResource ResourceName}`.

---

## 4. Agenda Engine Compatibility Strategy (EE4P & Think-Cell)

**Goal:** Ensure the native oPenEfficiency Agenda Wizard can safely coexist with or integrate with third-party formats.

### Efficient Elements (EE4P) Integration
- **Adapter Pattern:** The `AgendaGenerationService` must be able to serialize our lightweight `AgendaItemData` and native PowerPoint Sections into the specific XML schema expected by EE4P.
- **Tag Injection:** When generating shapes, inject the corresponding `EE4P_AGENDAWIZARD` string tags (e.g., `Topic`, `TimeSlot`) and use their GUID-based naming patterns (e.g., `ee4p_item_[guid]_shape`).

### ThinkCellGuard
- **Safe Coexistence (No Overwrite):** Full compatibility is not feasible due to Think-Cell's proprietary hashing and internal rendering engine.
- **ThinkCellGuard:** The `AgendaService` must actively scan the presentation for `THINKCELLPRESENTATIONDONOTDELETE` and `THINKCELLSHAPEDONOTDELETE` tags. Any slide containing these tags must be skipped and locked from automated oPenEfficiency updates to prevent corruption of the user's Think-Cell links.

---

## 5. Temporary Files & Workspace Hygiene

**Mandate:** The project root MUST remain clean at all times. AI agents MUST store all custom scripts (Python, PowerShell, Bash), temporary logs, and intermediate outputs EXCLUSIVELY in the **`_temp_scripts/`** directory.
- **Root Cleanup:** Do not create new scripts, `.txt`, or `.ps1` files directly in the root directory. 
- **Naming:** Use descriptive names for scripts (e.g., `_temp_scripts/update_sidebar_layout.py`) to maintain traceability.
- **Cleanup:** Delete temporary scripts once the task is fully verified.

---

## 6. Communication, Uncertainty & Documentation

- **Uncertainty:** If you are unsure about any implementation detail, architectural decision, or the exact scope of a requested change, you **MUST** ask the user for clarification before proceeding with file modifications. It is better to confirm intent than to guess and introduce technical debt.
- **Knowledge Persistence:** Whenever a complex bug, compilation error, or architectural mismatch is resolved—especially those specific to VSTO/WPF integration—you **MUST** document the root cause and the permanent fix in `.agents/skills/vsto-pitfalls.md`.

---

## 7. Versioning Policy

**Automated Naming:** The application version is automatically updated during the GitHub Actions `publish` workflow. The format is **`yyyy.mm.dd.RunNumber`**.
- **No Manual Action Required:** You do NOT need to manually update `<ApplicationVersion>` in `oPenEfficiency.csproj` or `AssemblyVersion` in `Properties\AssemblyInfo.cs`.
- **Auto-Increment:** Ensure `<AutoIncrementApplicationRevision>` is always set to `false` in the `.csproj` to prevent Visual Studio from overriding this scheme locally.
- **UI Version Update (Pre-Push):** Before performing a `git push`, you **MUST** manually update the hardcoded version string in the "About" tab of `UI/Dialogs/SettingsWindow.xaml` (e.g., `<TextBlock Text="Version 2026.05.18.xxx" ... />`) to reflect the current deployment version or date.

---

## 8. Testing Requirements

Before considering a task complete:
- [ ] Code compiles without errors (Check `.csproj` globbing if errors occur)
- [ ] Feature appears in UI/sidebar (if applicable)
- [ ] DetailedHelpText is added to FeatureMetadata
- [ ] Selection-based enabling works
- [ ] Errors are logged via `ExceptionLogger.Log()`
- [ ] UI follows dark theme and DPI-aware positioning
- [ ] No hardcoded colors (use StaticResource)