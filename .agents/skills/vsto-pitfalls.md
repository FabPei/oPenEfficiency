---
name: vsto-pitfalls
description: VSTO development, .csproj management, WPF lifecycle, XAML identity collisions, connectionId errors, WinForms integration, lifecycle issues, PowerPoint add-in bugs
---

# VSTO & WPF Integration Pitfalls for oPenEfficiency

This guide documents critical technical mismatches between modern development practices and the legacy **VSTO (Visual Studio Tools for Office)** architecture to prevent recurring build and runtime errors.

---

## 21. Mouse Hit-Test Offsets During/After Resizing

WPF controls hosted in `ElementHost` (WinForms) within a TaskPane can experience "drifting" hit-tests (mouse clicks/hovers appearing offset by a few pixels) after frequent resizing.

### The Problem: Visual/Host Desync
- **SYMPTOM:** Mouse hovers or clicks activate elements shifted by 5-10 pixels (often matching the sum of root margins/paddings).
- **REASON:** The `ElementHost` window and the WPF visual root can get out of sync regarding their coordinate systems during rapid WinForms layout updates. Sub-pixel rendering can also contribute to accumulation errors.

### The Fix: Explicit Invalidation & Rounding
1. **Enable Layout Rounding:** Set `UseLayoutRounding="True"` on the root `UserControl`. This forces WPF to align elements to physical pixels, preventing "half-pixel" drift.
2. **Force Invalidation on Resize:** In the `SizeChanged` event of the WPF control, call `this.InvalidateVisual();`.
3. **Post-Debounce Sync:** If using a timer to debounce UI rebuilds (common in sidebars), call `this.UpdateLayout();` and `this.InvalidateVisual();` even if a full rebuild is not performed.
4. **Padding Stability:** Prefer using `Margin` on child elements rather than `Padding` on the root `ScrollViewer` or `UserControl`, as margins are generally more stable for coordinate mapping at the host boundary.

---

## 1. Project File Management (.csproj)
... (rest of the file)

VSTO projects are strictly legacy-formatted and have a nuanced relationship with "globbing."

### The Rule: Selective Globbing
- **FEATURES FOLDER:** This project **USES globbing** for the `Features` folder: `<Compile Include="Features\**\*.cs" />`. 
  - **Reason:** There are over 150 feature files; explicit listing is prone to human error. If you remove this wildcard, hundreds of "Type not found" (CS0246) errors will occur.
- **UI FOLDER:** You **MUST NOT** use globbing for XAML-linked files. Add every `.cs` and `.xaml` in the `UI` folders **explicitly**.
  - **Reason:** Wildcards break the `DependentUpon` link. This results in `InitializeComponent` missing errors.
- **DEDUPLICATION:** NEVER list a file explicitly if it is already covered by a globbing pattern. This causes `CS2002` (File specified multiple times). 
- **MODEL EXTRACTION:** When moving shared models, ensure the new file is explicitly included in the `.csproj` if it falls outside a globbed folder.

---

## 2. XAML Identity & ConnectionId Collisions

WPF generates internal integer IDs (`connectionId`) for interactive elements during compilation.

### The Problem: Shared Namespaces
- **PITFALL:** Nesting complex `ContextMenu` or `Resources` directly inside `ListBox` or `DataTemplate` in shared user controls.
- **SYMPTOM:** Runtime exception: *"Durch das Festlegen von connectionId wurde eine Ausnahme ausgelöst"* (Exception triggered by setting connectionId).
- **FIX:** Define `ContextMenu` and common styles as **StaticResources** at the top of the file and reference them by key.

---

## 3. Lifecycle: WPF-in-WinForms Integration

oPenEfficiency hosts WPF controls inside a WinForms `ElementHost` within a PowerPoint TaskPane.

### The Rule: Prefer Constructor over 'Loaded' Event
- **PITFALL:** Relying on the `UserControl_Loaded` event to initialize or populate UI data.
- **SYMPTOM:** The window appears completely **blank** or white.
- **REASON:** In VSTO, the `Loaded` event is unreliable and may not fire if the TaskPane is hidden/shown rapidly.
- **FIX:** Instantiate child controls and set default selections directly in the **Constructor** (after `InitializeComponent`).

---

## 4. Layout Constraints in TaskPanes

PowerPoint TaskPanes have unique resizing behaviors that can break standard WPF layouts.

### The Rule: Explicit Scrolling Constraints
- **PITFALL:** Omitting `HorizontalScrollBarVisibility="Disabled"` in `ListBox` or `WrapPanel`.
- **SYMPTOM:** Items display in one infinite horizontal row instead of wrapping into a grid.
- **FIX:** Always disable horizontal scrolling for gallery-style views to force vertical wrapping.

---

## 5. COM Interop Reliability

### Office Namespace Prefixes
- **PITFALL:** Using generic names like `SmartArtNode` or `Shape`.
- **SYMPTOM:** `CS0246` (Type or namespace name not found).
- **FIX:** Use fully qualified names (e.g., `Microsoft.Office.Core.SmartArtNode`) to distinguish between standard .NET types and PowerPoint COM types.

---

## 6. XAML Syntax Precision

Modern WPF compilers or specialized VSTO build tasks can be extremely sensitive to shorthand XAML syntax.

### The Rule: Proper Setter Syntax
- **PITFALL:** Using attribute-style shorthand in Setters, especially when using `TargetName`.
- **SYMPTOM:** `MC3072` (Property 'X' does not exist in XML namespace).
- **DO:** `<Setter TargetName="Bd" Property="Background" Value="Red" />`
- **Why:** The `Setter` class in WPF requires the `Property` and `Value` attributes to be explicitly defined.

---

## 7. XAML Encoding & Special Characters

VSTO build tasks often fail to handle non-ASCII characters correctly if the file encoding is inconsistent.

### The Rule: ASCII or Hex-Code Only
- **PITFALL:** Using special characters like `®`, `⚡`, or `📁` directly in XAML text.
- **SYMPTOM:** `CS0103` (The name 'ControlName' does not exist in context). The compiler fails to generate the `.g.cs` file correctly.
- **FIX:** Use XML entities (e.g., `&#174;` for `®`). Ensure the file is saved as **UTF-8 with BOM**.

---

## 8. .csproj XML Structural Integrity

VSTO projects require explicit file listing, making manual edits to the `.csproj` file common but dangerous.

### The Rule: Balanced Tags & Valid Schema
- **PITFALL:** Accidentally closing a `<Compile>` tag with a `</Page>` tag.
- **SYMPTOM:** `MSB4025` (The project file could not be loaded).
- **FIX:** Always ensure tag pairs match. Visual Studio aggressively locks the `.csproj` file; close it before saving structural changes.

---

## 9. VSToolsPath & Office Build Targets

The build system depends on specialized MSBuild targets.

### The Rule: Proper Tools Versioning
- **PITFALL:** Incorrect `VisualStudioVersion` or `VSToolsPath` condition logic.
- **SYMPTOM:** `MSB4019` (The imported project ... targets was not found).
- **FIX:** Ensure `VisualStudioVersion` is set to `17.0` (for VS 2022). The `VSToolsPath` condition must check for empty strings (`== ''`).

---

## 10. Destructive Refactoring & Misdiagnosing Build Errors

Massive structural changes or UI updates can cause misleading build errors.

### The Rule: Do Not Purge Code to Fix Build Errors
- **PITFALL:** Encountering `InitializeComponent` missing errors and assuming the C# logic is broken.
- **FIX:** Never mix structural `.csproj` changes with massive C# logic refactoring. Check for globbing in `.csproj` or XAML encoding issues first.

---

## 11. XAML-Logic Synchronization During Refactoring

Refactoring data models (e.g., renaming classes) often breaks XAML templates.

### The Rule: Update 'DataType' in XAML
- **PITFALL:** Renaming model classes in C# but leaving the old class names in XAML `DataType`.
- **SYMPTOM:** `CS0246` or empty UI rendering.
- **FIX:** Perform a project-wide search for the old class name in `.xaml` files and update the `DataType`.

---

## 12. WPF Drag & Drop vs. Interactive Elements

Overlapping hit-test logic can break interactive child elements.

### The Rule: Do Not Disable Hit Testing on Actionable Buttons
- **PITFALL:** Setting `IsHitTestVisible="False"` on a `Button` to allow the parent to capture drag events.
- **SYMPTOM:** The button becomes unclickable.
- **FIX:** Handle drag initiation via `PreviewMouseLeftButtonDown` on the container and ignore it if the source is an interactive control.

---

## 13. Namespace Resolution Errors (CS0234) for UI Dialogs

UI components often don't follow physical folder namespaces.

### The Rule: Check the XAML.cs Namespace
- **PITFALL:** Assuming `UI\Dialogs\MyDialog.xaml.cs` uses the `oPenEfficiency.UI.Dialogs` namespace.
- **FIX:** Most dialogs in this project use the root `oPenEfficiency.UI` namespace. Always check the file's `namespace` declaration.

---

## 14. WPF Text Input Events & VSTO Selection Desync

Rapid events can cause PowerPoint's selection context to become unreliable.

### The Rule: Cache the Target Shape
- **PITFALL:** Relying on `manager.GetSelectedShapes()` during rapid events (e.g., `TextChanged`).
- **FIX:** Cache the `_targetShape` reference when the toolbar opens and use boolean guards (`_isUpdating`) to prevent concurrent updates.

---

## 15. AlternativeText for Shape Identification

The project relies on `AlternativeText` to identify interactive shapes.

### The Rule: Always Tag New Groups Immediately
- **PITFALL:** Setting `Tags` but forgetting `AlternativeText`.
- **SYMPTOM:** Shape renders but clicking it does nothing.
- **FIX:** Explicitly set `group.AlternativeText = "your_tag_string"` immediately after grouping.

---

## 16. Missing Namespace Aliases (CS0103 / CS0246)

### The Rule: Always Alias Office Core
- **FIX:** Always include `using Office = Microsoft.Office.Core;` and `using PowerPoint = Microsoft.Office.Interop.PowerPoint;` when interacting with native PowerPoint enums.

---

## 17. Redundant DragMove() Calls and Event Bubbling

Registering `DragMove()` in multiple places causes crashes.

### The Rule: Don't Register DragMove on Both Window and Header
- **PITFALL:** Adding `MouseDown="Window_MouseDown"` to both the `<Window>` tag and a child title bar.
- **SYMPTOM:** `System.InvalidOperationException: DragMove can only be called when the primary mouse button is down.`
- **REASON:** WPF events bubble up, causing double-invocation.
- **FIX:** Register the handler **only** on the `<Window>` tag. Ensure the header Grid has `Background="Transparent"` to allow bubbling.

---

## 18. Locale-Specific XAML Parsing for Numeric Attributes

XAML parsing can conflict with regional decimal separators.

### The Rule: Prefer Space Separators in Multi-Value Attributes
- **PITFALL:** Using `RenderTransformOrigin="0.5,0.5"` with a comma.
- **SYMPTOM:** `XamlParseException` on systems where the comma is a decimal separator (e.g., Germany).
- **FIX:** Use a space separator: `RenderTransformOrigin="0.5 0.5"`.

---

## 19. Defensive Clamping for COM Property Assignments

Invalid property values cause generic PowerPoint crashes.

### The Rule: Clamp Values Before Assignment
- **FIX:** 
  1. Clamp coordinates to slide boundaries.
  2. Ensure `FontSize` is at least 1.
  3. Validate that margins are strictly less than shape dimensions.
  4. Ensure Enum values are within their 1-indexed range (not 0).
- **DEBUGGING:** Wrap assignments in surgical try-catch blocks that report the specific property and value.

---

## 20. Synchronized Layout Paths & Subdirectory Handling

Data consistency across windows depends on unified path resolution.

### The Rule: Unified Path Resolution via JsonService
- **PITFALL:** Using inconsistent paths (absolute vs relative, root vs subfolder) for the same data.
- **SYMPTOM:** Data saved in one window doesn't appear in another.
- **FIX:** Always use `JsonService` with a **unified relative path string** (e.g., `@"AgendaLayouts\agenda_layouts.json"`) across all components.

## 22. Exception Handling & UI Dispatching

VSTO Add-ins run inside the PowerPoint host process. Unhandled exceptions or improperly threaded UI calls can instantly crash the entire application.

### The Problem: Cross-Thread UI Crashes
- **PITFALL:** Attempting to open a WPF Window (like an error dialog or progress bar) directly from a background thread or a raw COM event handler without using the Dispatcher.
- **SYMPTOM:** `System.InvalidOperationException: The calling thread must be STA, because many UI components require this.` Or a silent, hard crash of PowerPoint.

### The Rule: Use ExceptionLogger and Dispatcher
1. **Always Catch:** Every feature's `Execute` method must be wrapped in a `try-catch` block.
2. **Central Logging:** Use `ExceptionLogger.Log(ex, "Feature.Context", showErrorUI: true)` for user-facing actions.
3. **Dispatcher Safety:** If you must show custom UI from a background task, always wrap it:
   ```csharp
   System.Windows.Application.Current.Dispatcher.Invoke(() => {
       var myWindow = new CustomWindow();
       myWindow.ShowDialog();
   });
   ```

---

## Summary Checklist for VSTO Stability
1. [ ] Is every new file explicitly listed in `.csproj`? (Except Features folder)
2. [ ] Are all `xaml.cs` files correctly linked with `DependentUpon`?
3. [ ] Are context menus defined as `StaticResources`?
4. [ ] Is the primary UI population logic in the Constructor?
5. [ ] Is horizontal scrolling disabled for grid-based views?
6. [ ] Are all XML tags in `.csproj` perfectly balanced?
7. [ ] Are non-ASCII characters in XAML replaced with XML entities?
8. [ ] Is `DragMove()` registered only at the Window level?
9. [ ] Are `RenderTransformOrigin` values space-separated (e.g., "0.5 0.5")?
10. [ ] Are shape dimensions and margins clamped to valid PowerPoint ranges?
11. [ ] Do all windows use the same relative path for shared JSON configuration?
