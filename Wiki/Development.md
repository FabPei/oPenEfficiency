# Development Guide

oPenEfficiency is built on the .NET Framework 4.8 using Microsoft's VSTO (Visual Studio Tools for Office).

## Project Structure

```text
oPenEfficiency/
├── Features/           # Core feature logic (Alignment, Tables, Visuals, etc.)
├── Models/             # Data structures and configuration models
├── Services/           # Business logic and PowerPoint interaction layer
├── UI/                 # WPF dialogs, sidebar panels, and custom controls
└── Utils/              # Helper classes and Win32 hooks
```

## Architectural Guidelines (from GEMINI.md)

When contributing or developing new features, strictly adhere to the following mandates:

### Technical Integrity & VSTO Stability

1. **NO GLOBBING in `.csproj`:** NEVER use wildcards like `**\*.cs` or `**\*.xaml` for UI, Models, Services, or Utils folders. Every file MUST be explicitly listed with its correct `DependentUpon` mapping. Wildcards break VSTO code-behind linking. The ONLY exception permitted is the `Features` folder.
2. **XAML connectionId Safety:** Define all complex `ContextMenu` and `Resources` as **StaticResources** at the top of the XAML file. NEVER nest anonymous context menus inside ListBox templates.
3. **Constructor-First Initialization:** Perform all primary UI population and child control instantiation in the **Constructor**, not the `Loaded` event. The `Loaded` event is suppressed in VSTO's `ElementHost`.
4. **Scrolling Constraints:** Gallery views (ListBox, WrapPanel) MUST have `HorizontalScrollBarVisibility="Disabled"` to force grid wrapping in the PowerPoint TaskPane.
5. **Office Namespace Precision:** Always use fully qualified names for Office COM types (e.g., `Microsoft.Office.Core.SmartArtNode`).
6. **Strict XAML Setter Syntax:** ALWAYS use the explicit `<Setter Property="X" Value="Y" />` syntax. NEVER use shorthand attributes like `<Setter X="Y" />`.

### Agenda Engine Compatibility Strategy

- **Efficient Elements (EE4P) Integration:** Use the `AgendaGenerationService` adapter pattern to serialize data for EE4P schemas and inject `EE4P_AGENDAWIZARD` tags.
- **ThinkCellGuard:** The `AgendaService` actively scans for `THINKCELLPRESENTATIONDONOTDELETE` and `THINKCELLSHAPEDONOTDELETE` tags. Slides containing these must be skipped to prevent corruption.

### UI/UX & Theming Guidelines

- **Theme Cohesion:** Never use `TextMutedBrush` or `TextSecondaryBrush` for standalone icon buttons or close ("X") buttons in floating windows in Light Theme. Use `TextPrimaryBrush`.
- **Tooltip Inheritance:** WPF ToolTips natively inherit `Foreground` properties. Always ensure `ToolTip` styles explicitly set their own `Background` and `Foreground` decoupled from the button hover states.