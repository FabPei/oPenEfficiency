# oPenEfficiency Project Skills & Knowledge

This document captures key architectural decisions, patterns, and project structure for future AI agents working on this codebase.

## Project Overview

**oPenEfficiency** is a VSTO Add-in for Microsoft PowerPoint that provides enhanced productivity features including:
- Infographic shapes (Harvey Balls, Star Ratings, Thermometers, Traffic Lights, Checkboxes)
- Shape alignment, matching, and arrangement tools
- Table productivity tools (Transpose, Sort, Sum, Split)
- Agenda generation and management (Agenda Wizard)
- Asset Library for slides, shapes, and images
- Custom sidebar and floating toolbars

**Technology Stack:**
- .NET Framework 4.8
- VSTO (Visual Studio Tools for Office)
- WPF for UI dialogs and controls
- Win32 API hooks for global mouse capture
- COM Interop for PowerPoint automation

---

## Project Structure

```
oPenEfficiency/
├── Services/                    # Core service layer
│   ├── PowerPointManager.cs     # Central PowerPoint access point
│   ├── ShortcutManager.cs       # Global hotkey management
│   └── AgendaService.cs         # Agenda/CustomXML data persistence
│
├── Features/                    # Feature implementations (organized by category)
│   ├── Alignment/               # Shape alignment and positioning
│   ├── Charts/                  # Infographic features (HarveyBall, StarRating, etc.)
│   ├── Tables/                  # Table utilities
│   ├── Text/                    # Text manipulation features
│   ├── Utilities/               # General utilities (Cleaner, StickyNotes, etc.)
│   └── Visuals/                 # Visual enhancements (Checkbox, ColorPicker, etc.)
│
├── UI/                          # User interface components
│   ├── Panels/                  # Sidebar and task pane controls
│   ├── Dialogs/                 # Modal dialogs (Settings, Wizard dialogs)
│   ├── Toolbars/                # Floating toolbars
│   ├── Menus/                   # Context menus and floating menus
│   ├── Wizards/                 # Multi-step wizards (Agenda, Export)
│   ├── Controls/                # Reusable WPF controls
│   └── Converters/              # WPF value converters
│
├── Utils/                       # Utility classes
│   ├── ExceptionLogger.cs       # Centralized exception logging
│   ├── Win32MouseHook.cs        # Global mouse hook for drag operations
│   ├── JsonService.cs           # JSON serialization helpers
│   └── ClipboardHelpers.cs      # Clipboard operations
│
├── Models/                      # Data models
│   ├── AgendaLayoutConfig.cs    # Agenda configuration
│   └── ...
│
├── ThisAddIn.cs                 # VSTO entry point and ribbon callbacks
└── oPenEfficiency.csproj        # Project file (VSTO-specific imports)
```

---

## Key Architectural Patterns

### 1. Centralized PowerPoint Access
All PowerPoint COM operations flow through `PowerPointManager`:
```csharp
var manager = new PowerPointManager(this.Application);
var shapes = manager.GetSelectedShapes();
var slide = manager.GetCurrentSlide();
```

**Why:** Isolates COM interop, provides consistent error handling, simplifies testing.

### 2. Feature-Based Organization
Features are organized by **category** (Alignment, Charts, Tables) not by **type** (Models, Views, Controllers).

**Why:** Makes it easy to find all code related to a specific feature. Adding a new alignment feature? Look in `Features/Alignment/`.

### 3. Static Feature Classes
Most features are `public static class` with an `Execute(PowerPointManager manager)` method:
```csharp
public static class HarveyBallFeature
{
    public static bool Execute(PowerPointManager manager, float percentage = 0.25f)
    {
        // Implementation
    }
}
```

**Why:** Features are stateless operations. Static classes make intent clear and avoid unnecessary instantiation.

### 4. Centralized Exception Logging
Use `ExceptionLogger.Log(ex, context)` instead of empty catch blocks:
```csharp
catch (Exception ex)
{
    ExceptionLogger.Log(ex, "ThisAddIn.ProcessSelectionChange");
    return null;
}
```

**Why:** Silent exceptions make debugging impossible. All errors are logged to `%TEMP%\oPenEfficiency\error_YYYYMMDD.log`.

### 5. Win32 Mouse Hook for Drag Operations
The add-in uses a global mouse hook (`Win32MouseHook`) to track dragging for features like Snap-to-Objects and Snap-to-Grid:
```csharp
_mouseHook = new Win32MouseHook();
_mouseHook.LeftButtonUp += OnGlobalMouseUp;
_mouseHook.MouseMoveFilter = OnMouseMoveFilter;
```

**Important:** Always call `Dispose()` to prevent system-wide hook leaks.

---

## Critical Implementation Details

### WPF Application Singleton
VSTO does not create a WPF `Application` singleton automatically. The add-in creates one in `ThisAddIn_Startup`:
```csharp
private static void EnsureWpfApplication()
{
    if (System.Windows.Application.Current == null)
    {
        var app = new System.Windows.Application();
        app.ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown;
    }
}
```

**Why:** WPF `Window.Show()` requires an `Application` instance. `OnExplicitShutdown` prevents the add-in from closing when dialogs close.

### DPI-Aware Positioning
Floating toolbars must account for system DPI:
```csharp
private float _dpiX = 96f, _dpiY = 96f;

// Capture DPI at startup
using (var g = Graphics.FromHwnd(IntPtr.Zero))
{
    _dpiX = g.DpiX;
    _dpiY = g.DpiY;
}

// Convert screen pixels to WPF units
toolbar.Left = (pixelX * 96.0 / _dpiX) + 10;
```

### Shape Tags for Metadata
PowerPoint shapes have a `Tags` collection for storing metadata:
```csharp
shape.Tags.Add("oPE_Widget", "HarveyBall");
shape.Tags.Add("Value", "0.75");
```

**Pattern:** Use `oPE_` prefix for all custom tags to avoid conflicts.

### AlternativeText for Feature State
Complex shapes store their state in `AlternativeText`:
```csharp
// Format: "Tag|value1|value2|..."
group.AlternativeText = $"{HarveyBallTag}|{percentage}";
group.AlternativeText = $"{StarRatingTag}|{count}|{rating}|{iconMode}|{colorBgr}";
```

**Why:** AlternativeText persists with the shape and survives copy/paste operations.

---

## Common Operations

### Get Selected Shapes
```csharp
var shapes = manager.GetSelectedShapes();
if (shapes != null && shapes.Count == 1)
{
    var shape = shapes[1]; // 1-indexed!
}
```

### Detect Feature Shapes
```csharp
if (shape.AlternativeText.StartsWith(HarveyBallFeature.HarveyBallTag))
{
    // Handle Harvey Ball
}
```

### Create Grouped Shapes
```csharp
var shapeNames = new string[] { bg.Name, pie.Name };
var group = slide.Shapes.Range(shapeNames).Group();
group.AlternativeText = $"{HarveyBallTag}|{percentage}";
group.Name = "HarveyBall_" + Guid.NewGuid().ToString().Substring(0, 8);
```

### Show Floating Toolbar
```csharp
var toolbar = new FloatingToolbar();
toolbar.AddLabel("Label");
toolbar.AddButton("Action", () => { /* handler */ });

var window = Application.ActiveWindow;
int pixelX = window.PointsToScreenPixelsX(shape.Left + shape.Width);
int pixelY = window.PointsToScreenPixelsY(shape.Top);

toolbar.Left = (pixelX * 96.0 / _dpiX) + 10;
toolbar.Top = (pixelY * 96.0 / _dpiY);
toolbar.Show();
```

---

## Build & Deployment

### Project File Requirements
The `.csproj` file imports VSTO-specific targets:
```xml
<Import Project="$(MSBuildExtensionsPath)\Microsoft\VisualStudio\v18.0\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets" />
```

**Note:** Building requires Visual Studio with Office Developer Tools. `dotnet build` alone will fail.

### Certificate Signing
The project uses a PFX certificate for ClickOnce manifests. Only `oPenEfficiency_TemporaryKey.pfx` is actively used.

### Git Ignore Patterns
Build artifacts excluded in `.gitignore`:
```
*.dll, *.exe, *.pdb, *.vsto, *.zip, *.log
obj/, Publish/, Application Files/
```

---

## Debugging Tips

### Enable VSTO Debug Output
1. In Visual Studio: Project Properties → Debug → Check "Enable VSTO Debugging"
2. Set breakpoints in `ThisAddIn_Startup` for initialization issues

### Exception Logs
Runtime errors are logged to:
```
%TEMP%\oPenEfficiency\error_YYYYMMDD.log
```

### Shortcut Debug Log
Hotkey registration issues are logged to:
```
%TEMP%\oPenEfficiency\ShortcutDebug_oPenEfficiency.txt
```

### Mouse Hook Debugging
The `Win32MouseHook` class writes debug output on install/uninstall:
```csharp
Debug.WriteLine("Win32MouseHook successfully installed on thread {threadId}.");
```

---

## Known Limitations

1. **Single Instance WPF:** Multiple PowerPoint instances share a static WPF `Application`. This is by design but requires careful resource cleanup.

2. **COM Reference Leaks:** Always release COM objects explicitly:
   ```csharp
   var shapes = selection.ShapeRange;
   // Use shapes...
   Marshal.ReleaseComObject(shapes); // Optional, GC usually handles it
   ```

3. **Hotkey Conflicts:** Windows allows only one handler per hotkey combination. The app shows a warning if registration fails.

4. **Threading:** VSTO event handlers run on the PowerPoint UI thread. Long operations should use `Dispatcher.BeginInvoke` or `Task.Run`.

---

## Adding New Features

For a complete, step-by-step guide, see **[Adding Features](./skills/adding-features.md)**.

### Quick Reference (New Attribute-Based Pattern)

| Step | File | Action |
|------|------|--------|
| 1 | `Features/[Cat]/[Name].cs` | Create feature class with `[FeatureMetadata]` attribute |
| 2 | *(Optional)* | Add wrapper method if feature has parameters |
| 3 | `UI/Dialogs/` | Create dialog if feature requires user input (manual execution handling) |

**That's it!** Features with `[FeatureMetadata]` are auto-discovered via reflection. No manual registration needed.

### Exceptions (Manual Handling Required)

| Feature Type | Additional Steps |
|--------------|------------------|
| **Dialog features** | Create dialog, add execution case in `MainSidebar.xaml.cs` |
| **Toggle features** | Add execution case in `MainSidebar.xaml.cs` (ToggleButton handling) |
| **Features with parameters** | Add wrapper `Execute(manager)` with defaults, or manual handling |

### Basic Pattern (Auto-Discovered)
```csharp
// Features/YourCategory/YourFeature.cs
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.YourCategory
{
    [FeatureMetadata(
        Id = "BtnYourFeature",
        Name = "Your Feature",
        Tooltip = "Short tooltip",
        IconData = "M12,2L14,4...",
        Color = "#10B981",
        Description = "Full description",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class YourFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var shapes = manager.GetSelectedShapes();
                // Implementation
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "YourFeature.Execute");
                return false;
            }
        }
    }
}
```

### Toggle Feature Pattern
For features with on/off state (e.g., snap-to-grid):

**Note:** Toggle features require manual execution handling in `MainSidebar.xaml.cs`:

```csharp
// Features/Utilities/SnapToGridFeature.cs
[FeatureMetadata(
    Id = "BtnSnapToGrid", Name = "Snap to Grid", ..., IsToggle = true)]
public static class SnapToGridFeature
{
    public static bool Execute(PowerPointManager manager)
    {
        // Return false to trigger manual switch handling
        return false;
    }

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

**Manual execution in `MainSidebar.xaml.cs`:**
```csharp
case "BtnSnapToGrid":
    var toggleBtn = targetBtn as ToggleButton;
    SnapToGridFeature.Execute(manager, toggleBtn);
    break;
```

---

## References

- [VSTO Architecture](https://docs.microsoft.com/en-us/visualstudio/vsto/vsto-architecture-overview)
- [PowerPoint Interop](https://docs.microsoft.com/en-us/office/vba/api/powerpoint.publisher)
- [WPF in VSTO](https://docs.microsoft.com/en-us/visualstudio/vsto/wpf-in-vsto-add-ins)
