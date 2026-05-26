---
name: adding-features
description: Add new features using attribute-based auto-discovery, [FeatureMetadata] attribute, simplified registration, sidebar integration, toggle buttons, wrapper methods for parameters
---

# Adding Features Guide (Attribute-Based Auto-Discovery)

This guide shows how to add new features to the oPenEfficiency PowerPoint Add-in using the **attribute-based auto-discovery** system.

## Overview

**Before (Legacy - 7 steps):**
1. Create feature class
2. Create service class (if stateful)
3. Register service in ThisAddIn.cs
4. Add to `FeatureLibrary.AllFeatures` list
5. Add to `FeatureLibrary.GetFeatureInfo()` switch
6. Add execution case in `MainSidebar.xaml.cs`
7. Add to `.csproj`

**Now (Attribute-Based - 1-2 files):**
1. Create feature class with `[FeatureMetadata]` attribute
2. *(Optional)* Add wrapper method if feature has parameters or opens dialogs

**Benefits:**
- No manual registration in FeatureLibrary.cs
- No manual registration in MainSidebar.xaml.cs (for simple features)
- Features are self-describing (metadata lives with the code)
- Easier for open-source contributors

---

## Step 1: Create the Feature Class

Create a new file in `Features/[Category]/`:

```csharp
// Features/Alignment/AlignCenterFeature.cs
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Alignment
{
    [FeatureMetadata(
        Id = "BtnAlignCenter",
        Name = "Align Center",
        Tooltip = "Align centers of selected shapes",
        IconData = "M12,2L14,4L10,8L12,10L16,6L20,10L22,8L18,4L22,2L20,0L16,4L12,0L10,2L14,6L10,10L12,12L16,8L20,12L18,14L14,10L10,14L8,12L12,8L8,4L6,6L10,10L6,14L8,16L12,12L16,16L18,14L14,10L18,6L20,8L16,12L20,16L18,18L14,14L10,18L12,20L16,16L20,20L18,22L14,18L10,22L8,20L12,16L8,12L6,14L10,18L6,22L8,24L12,20L16,24L18,22L14,18L18,14L20,16L16,20L20,24L18,26L14,22L10,26L8,24L12,20L8,16L6,18L10,22L6,26L8,28L12,24L16,28L18,26L14,22L18,18L20,20L16,24L20,28L18,30L14,26L10,30L8,28L12,24L8,20L6,22L10,26L6,30L8,32L12,28L16,32L18,30L14,26L18,22L20,24L16,28L20,32L18,34L14,30L10,34L8,32L12,28L8,24L6,26L10,30L6,34L8,36L12,32L16,36L18,34L14,30L18,26L20,28L16,32L20,36L18,38L14,34L10,38L8,36L12,32L8,28L6,30L10,34L6,38L8,40L12,36L16,40L18,38L14,34L18,30L20,32L16,36L20,40L18,42L14,38L10,42L8,40L12,36L8,32L6,34L10,38L6,42L8,44L12,40L16,44L18,42L14,38L18,34L20,36L16,40L20,44L18,46L14,42L10,46L8,44L12,40L8,36L6,38L10,42L6,46L8,48L12,44L16,48L18,46L14,42L18,38L20,40L16,44L20,48L18,50L14,46L10,50L8,48L12,44L8,40L6,42L10,46L6,50L8,52L12,48L16,52L18,50L14,46L18,42L20,44L16,48L20,52L18,54L14,50L10,54L8,52L12,48L8,44L6,46L10,50L6,54L8,56L12,52L16,56L18,54L14,50L18,46L20,48L16,52L20,56L18,58L14,54L10,58L8,56L12,52L8,48L6,50L10,54L6,58L8,60L12,56L16,60L18,58L14,54L18,50L20,52L16,56L20,60L18,62L14,58L10,62L8,60L12,56L8,52L6,54L10,58L6,62L8,64L12,60L16,64L18,62L14,58L18,54L20,56L16,60L20,64L18,66L14,62L10,66L8,64L12,60L8,56L6,58L10,62L6,66L8,68L12,64L16,68L18,66L14,62L18,58L20,60L16,64L20,68L18,70L14,66L10,70L8,68L12,64L8,60L6,62L10,66L6,70L8,72L12,68L16,72L18,70L14,66L18,62L20,64L16,68L20,72L18,74L14,70L10,74L8,72L12,68L8,64L6,66L10,70L6,74L8,76L12,72L16,76L18,74L14,70L18,66L20,68L16,72L20,76L18,78L14,74L10,78L8,76L12,72L8,68L6,70L10,74L6,78L8,80L12,76L16,80L18,77",
        Color = "#4CAF50",
        Description = "Aligns the horizontal centers of all selected shapes",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class AlignCenterFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var shapes = manager.GetSelectedShapes();
                if (shapes == null || shapes.Count < 2)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Please select at least 2 shapes.", "Info",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return false;
                }

                // Implementation: Align all shapes to the first shape's center
                var firstShape = shapes[1];
                float targetLeft = firstShape.Left + (firstShape.Width / 2);

                for (int i = 2; i <= shapes.Count; i++)
                {
                    var shape = shapes[i];
                    shape.Left = targetLeft - (shape.Width / 2);
                }

                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "AlignCenterFeature.Execute");
                System.Windows.Forms.MessageBox.Show(
                    $"Error: {ex.Message}", "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }
    }
}
```

### Attribute Properties

| Property | Type | Required | Description |
|----------|------|----------|-------------|
| `Id` | string | Yes | Button ID (must be unique, prefix with "Btn") |
| `Name` | string | Yes | Display name in sidebar |
| `Tooltip` | string | Yes | Hover text on button |
| `IconData` | string | Yes | SVG path data for icon |
| `Color` | string | Yes | Hex color for button styling |
| `Description` | string | Yes | Short description for tooltip/settings |
| `DetailedHelpText` | string | Yes | Comprehensive explanation for the Feature Explorer Wiki. MUST document any right-click Context Menu options here (e.g. "Right-click options:\n• Do something"). |
| `HelpImagePath` | string | No | Path to animated GIF/PNG for Feature Explorer (e.g. "pack://application:,,,/oPenEfficiency;component/UI/Assets/Help/MyFeature.gif") |
| `MinSelection` | int | No | Minimum shapes required (default: 0) |
| `MaxSelection` | int | No | Maximum shapes (default: 0 = unlimited) |
| `RequiredType` | PpSelectionType | No | Required selection type (default: ppSelectionShapes) |
| `RequiresTable` | bool | No | Must have table selected (default: false) |
| `IsToggle` | bool | No | Use ToggleButton instead of Button (default: false) |

---

## Step 2: Handle Special Cases

### Case A: Feature with Parameters (Add Wrapper)

If your feature has parameters, add a wrapper method with sensible defaults:

```csharp
// Features/Charts/HarveyBallFeature.cs
[FeatureMetadata(
    Id = "BtnHarveyBall", Name = "Harvey Ball", ..., MinSelection = 0)]
public static class HarveyBallFeature
{
    /// <summary>
    /// Wrapper for auto-discovery - creates Harvey Ball with default 25%.
    /// </summary>
    public static bool Execute(PowerPointManager manager)
    {
        return Execute(manager, 0.25f);
    }

    /// <summary>
    /// Main implementation with percentage parameter.
    /// </summary>
    public static bool Execute(PowerPointManager manager, float percentage)
    {
        // Implementation with parameter
        CreateHarveyBall(manager, percentage);
        return true;
    }
}
```

### Case B: Feature Opens Dialog (Return false)

If your feature opens a dialog, add the attribute but return `false` to trigger manual handling:

```csharp
// Features/Utilities/GanttChartFeature.cs
[FeatureMetadata(
    Id = "BtnGanttChart", Name = "Gantt Chart", ..., MinSelection = 0)]
public static class GanttChartFeature
{
    /// <summary>
    /// Wrapper for auto-discovery - returns false to open dialog.
    /// </summary>
    public static bool Execute(PowerPointManager manager)
    {
        // Return false to trigger manual switch handling for dialog
        return false;
    }

    /// <summary>
    /// Main implementation - called from manual switch after dialog.
    /// </summary>
    public static bool Create(PowerPointManager manager, DateTime startDate, DateTime endDate)
    {
        // Implementation after dialog configuration
        return true;
    }
}
```

**Add manual execution in `MainSidebar.xaml.cs`:**
```csharp
case "BtnGanttChart":
    var dialog = new GanttChartDialog();
    if (dialog.ShowDialog() == true)
    {
        GanttChartFeature.Create(manager, dialog.StartDate, dialog.EndDate);
    }
    break;
```

### Case C: Toggle Feature (Manual Handling)

Toggle features (SnapToGrid, SnapToObjects, FlightMode) require ToggleButton handling:

```csharp
// Features/Utilities/SnapToGridFeature.cs
[FeatureMetadata(
    Id = "BtnSnapToGrid", Name = "Snap to Grid", ..., IsToggle = true)]
public static class SnapToGridFeature
{
    /// <summary>
    /// Wrapper for auto-discovery - returns false for manual toggle handling.
    /// </summary>
    public static bool Execute(PowerPointManager manager)
    {
        return false;
    }

    /// <summary>
    /// Main implementation with ToggleButton parameter.
    /// </summary>
    public static void Execute(PowerPointManager manager, ToggleButton toggleBtn = null)
    {
        var snapManager = Globals.ThisAddIn.SnapGridManager;
        snapManager.Toggle();
        if (toggleBtn != null)
        {
            toggleBtn.IsChecked = snapManager.IsEnabled;
        }
    }
}
```

**Add manual execution in `MainSidebar.xaml.cs`:**
```csharp
case "BtnSnapToGrid":
    var toggleBtn = targetBtn as ToggleButton;
    SnapToGridFeature.Execute(manager, toggleBtn);
    break;
```

---

## Step 3: Build and Verify

That's it! Build the project and verify:

```bash
# Build in Visual Studio
msbuild oPenEfficiency/oPenEfficiency.slnx /p:Configuration=Release

# Or press Ctrl+Shift+B in Visual Studio
```

**Checklist:**
- [ ] Code compiles without errors
- [ ] Feature appears in sidebar (correct section, icon, color)
- [ ] Tooltip displays on hover
- [ ] Selection-based enabling works (if MinSelection/RequiredType set)
- [ ] Feature executes without crashing
- [ ] Errors are logged via `ExceptionLogger.Log()`

---

## Step 4: UI/UX & Theming Requirements (CRUCIAL)

When adding features that include UI elements (like floating toolbars, dialogs, or settings panes), you **MUST** ensure they pass WCAG AA accessibility standards across both Light and Dark themes.

**Mandatory Rules:**
1. **Never Hardcode Colors:** Do not use hardcoded hex values (e.g., `#888888`, `#0078D4`) for foregrounds or backgrounds in XAML.
2. **Use DynamicResources:** Always bind colors to the global theme dictionary using `{DynamicResource [BrushName]}` (e.g., `{DynamicResource TextPrimaryBrush}`, `{DynamicResource SurfaceBrush}`).
3. **Contrast Testing:** Small interactive elements (like close "X" buttons or standalone icons) must have a contrast ratio of at least 4.5:1 against their background. If a button rests on a light `SurfaceBrush`, using a muted text brush (like `TextMutedBrush` or `TextSecondaryBrush`) will often fail contrast checks. Use `TextPrimaryBrush` or dedicated accent brushes (like `AccentTextBrush`) instead.
4. **Tooltip Inheritance:** WPF ToolTips natively inherit `Foreground` properties from their parent controls. If a parent button changes its text color on hover (e.g., to white), the tooltip will inherit it and become invisible on a light background. Always explicitly set `Background` and `Foreground` on `ToolTip` styles to decouple them from button hover states.

---

## Feature Patterns

### Pattern Comparison

| Pattern | Use Case | Example | Files Needed |
|---------|----------|---------|--------------|
| **Simple (auto)** | Single-action, no state | `AlignCenterFeature` | 1 (feature class) |
| **With parameters** | Takes configuration | `HarveyBallFeature` | 1 (wrapper in same file) |
| **Dialog (manual)** | Opens dialog/window | `GanttChartFeature` | 2 (feature + dialog) |
| **Toggle (manual)** | On/off state | `SnapToGridFeature` | 1 (manual execution) |
| **Service + toggle** | Stateful, hooks | `SnapToObjectsFeature` | 2 (feature + manager) |

---

## Selection-Based Enabling

Features automatically enable/disable based on selection:

```csharp
[FeatureMetadata(
    Id = "BtnAlignLeft",
    Name = "Align Left",
    MinSelection = 2,
    RequiredType = PpSelectionType.ppSelectionShapes)]
```

The sidebar's `OnSelectionChange()` method automatically updates button states.

### Selection Type Values

```csharp
PpSelectionType.ppSelectionNone      // No selection
PpSelectionType.ppSelectionShapes    // Shape selection
PpSelectionType.ppSelectionText      // Text selection
PpSelectionType.ppSelectionSlides    // Slide thumbnails
```

---

## Common Operations

### Get Selected Shapes
```csharp
var shapes = manager.GetSelectedShapes();
if (shapes != null && shapes.Count >= 2)
{
    var firstShape = shapes[1]; // 1-indexed!
}
```

### Create Shape
```csharp
var shape = slide.Shapes.AddShape(
    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
    left: 100, top: 100, width: 200, height: 100);
shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Colors.Blue);
```

### Create Group
```csharp
var shapeNames = new string[] { shape1.Name, shape2.Name };
var group = slide.Shapes.Range(shapeNames).Group();
group.AlternativeText = "oPE_Widget|YourFeature";
```

### Store Metadata
```csharp
// Using Tags (persists with shape)
shape.Tags.Add("oPE_Widget", "YourFeature");
shape.Tags.Add("FeatureId", Guid.NewGuid().ToString());

// Using AlternativeText (for state)
group.AlternativeText = $"YourFeatureTag|{value1}|{value2}";
```

### Get Metadata
```csharp
string widgetTag = GetTagValue(shape, "oPE_Widget");
if (widgetTag == "YourFeature")
{
    // Handle your feature shape
}
```

---

## Debugging Tips

### Feature Not Appearing in Sidebar
1. Check `[FeatureMetadata]` attribute is present and valid
2. Check feature class is `public static class`
3. Check `Execute(PowerPointManager)` method exists and returns `bool`
4. Verify `.csproj` includes glob pattern `<Compile Include="Features\**\*.cs" />`

### Feature Not Executing
1. Set breakpoint in `Execute` method
2. Check for exceptions in `%TEMP%\oPenEfficiency\error_YYYYMMDD.log`
3. For dialog features, verify manual switch case exists in `MainSidebar.xaml.cs`

### Selection-Based Enabling Not Working
1. Check `MinSelection`/`MaxSelection` values in attribute
2. Check `RequiredType` matches expected selection
3. Verify `OnSelectionChange()` is being called in `MainSidebar.xaml.cs`

---

## Legacy Pattern (Still Supported)

Features without `[FeatureMetadata]` still work via manual registration:

```csharp
// Old pattern - still functional for backward compatibility
// 1. Add to FeatureLibrary.AllFeatures
// NOTE: Always keep this list sorted alphabetically by Name!
new SidebarFeature { Id = "BtnLegacyFeature", Name = "Legacy" },

// 2. Add to FeatureLibrary.GetFeatureInfo()
case "BtnLegacyFeature":
    return new FeatureDisplayInfo { Tooltip = "...", Icon = "...", ... };

// 3. Add execution in MainSidebar.xaml.cs
case "BtnLegacyFeature":
    LegacyFeature.Execute(manager);
    break;
```

**IMPORTANT:** When adding a new feature (even via auto-discovery), ensure it is added to the `FeatureLibrary.AllFeatures` list in **alphabetical order by Name**. This ensures the "Available Features" list in the Settings menu remains organized.

**CRITICAL PITFALL - Duplicate Registrations & Missing Features:**
- **Missing Features:** Features that are auto-discovered via `[FeatureMetadata]` but *never added* to `FeatureLibrary.AllFeatures` will be completely missing from the Settings search pool and dynamic layouts like `GetAllFeaturesLayout()`. Always append new features to `FeatureLibrary.AllFeatures`.
- **Duplicate Features:** Do NOT add the same feature to `FeatureLibrary.AllFeatures` multiple times. Duplicate entries in this list will cause the feature to appear multiple times in the UI. Always search the list first to ensure the feature isn't already registered!
- **MANDATORY AI VALIDATION:** AI Agents are prone to "silent failures" when using text replacement tools on large lists like `AllFeatures`. After attempting to add a feature to `FeatureLibrary.AllFeatures`, the agent **MUST immediately use `grep_search` or `read_file`** to explicitly verify the string was injected correctly. Do not assume the edit succeeded.

**Recommendation:** Use attribute-based pattern for all new features. Keep legacy pattern only for:
- Features being migrated gradually
- Complex features with custom registration needs

---

## Quick Reference

### Simple Feature (1 file)
```csharp
[FeatureMetadata(Id = "BtnFeature", Name = "Feature", ...)]
public static class Feature
{
    public static bool Execute(PowerPointManager manager)
    {
        // Implementation
        return true;
    }
}
```

### Feature with Parameters (1 file, wrapper)
```csharp
[FeatureMetadata(Id = "BtnFeature", Name = "Feature", ...)]
public static class Feature
{
    public static bool Execute(PowerPointManager manager)
        => Execute(manager, defaultValue);

    public static bool Execute(PowerPointManager manager, string param)
    {
        // Implementation with parameter
        return true;
    }
}
```

### Dialog Feature (2 files, manual execution)
```csharp
[FeatureMetadata(Id = "BtnFeature", Name = "Feature", ...)]
public static class Feature
{
    public static bool Execute(PowerPointManager manager)
        => false; // Trigger manual dialog handling

    public static bool Create(PowerPointManager manager, Config config)
    {
        // Called from manual switch after dialog
        return true;
    }
}
```

### Toggle Feature (manual execution)
```csharp
[FeatureMetadata(Id = "BtnFeature", Name = "Feature", IsToggle = true)]
public static class Feature
{
    public static bool Execute(PowerPointManager manager)
        => false; // Trigger manual toggle handling

    public static void Execute(PowerPointManager manager, ToggleButton btn)
    {
        var service = Globals.ThisAddIn.ServiceManager;
        service.Toggle();
        btn.IsChecked = service.IsEnabled;
    }
}
```

## Common Pitfalls

### Feature Wrapper Bypassing UI (Hardcoded Defaults)
- **PITFALL:** Creating a wrapper `Execute(PowerPointManager manager)` method that skips the user interface (e.g., dialogs or toolbars) and directly inserts a shape with hardcoded default parameters.
- **SYMPTOM:** When the user clicks the feature button in the sidebar, no dialog opens, and a generic default object (e.g., a 2x2 grid or a 5-step progress bar) is inserted immediately.
- **FIX:** Ensure the parameterless `Execute` wrapper properly instantiates and `Show()`s the feature's configuration Window/Toolbar. The actual insertion logic should be triggered by the "Apply" or "Insert" buttons within that UI, not automatically upon clicking the sidebar button.

---

## Step 5: Documentation (MANDATORY)

Every feature implementation or modification is **INCOMPLETE** until you have updated the project documentation. Documentation is critical for feature discoverability and maintainability.

### 1. Update In-App "Wiki" (Metadata)
Ensure the `[FeatureMetadata]` attribute in your feature class has accurate and comprehensive text:
- **`Description`**: A clear, 1-2 sentence summary of what the feature does.
- **`DetailedHelpText`**: A detailed explanation of the feature's logic. You **MUST** document any right-click Context Menu options here (e.g., "Right-click options:\n• Top-Down: Sorts...").

### 2. Update Project History (Changelog)
Add an entry to `CHANGELOG.md` under the latest date heading. Be concise but descriptive.

### 3. Update Feature Discoverability (README)
Add the feature to the appropriate category in the `Features` section of `README.md`. Ensure the name matches what's in the Sidebar.

---

## Related Documentation & Context Compression
- [Workspace Hygiene](./workspace-hygiene.md) - Repository cleanup rules and changelog standards
- [VSTO Pitfalls](./vsto-pitfalls.md) - Handling COM objects and threading in WPF windows
- [UI Development Guide](./ui-development.md) - Creating dialogs and floating windows
- [Floating Window Creation](./floating-window-creation.md) - WPF window patterns
- [Project Overview](./project-overview.md) - Architecture and patterns
- [CLAUDE.md](../../CLAUDE.md) - Project-specific conventions and migration history
