---
name: ui-development
description: UI changes, WPF development, dialogs, XAML styling, sidebar modifications, control styling, theme colors, dark theme, UI architecture, modal dialogs, WPF windows, PowerPoint Add-in UI
---

# UI Changes Guide for oPenEfficiency

This guide helps AI agents make consistent UI changes in the oPenEfficiency PowerPoint Add-in.

---

## UI Architecture Overview

The UI layer follows a clear separation of concerns:

```
UI/
├── Panels/          # Docked task panes (sidebar, theme color pane)
├── Dialogs/         # Modal dialog windows (Settings, Wizards)
├── Toolbars/        # Floating toolbars (RatingToolbar, NumerationToolbar)
├── Menus/           # Context menus and floating menus (GanttChartMenuPanel)
├── Wizards/         # Multi-step wizards (ExportWizard, AgendaWizard)
├── Controls/        # Reusable WPF user controls
└── Converters/      # WPF value converters (FeatureIconConverter)
```

---

## Adding a New Dialog

### 1. Create XAML File
```xml
<!-- UI/Dialogs/YourDialog.xaml -->
<Window x:Class="oPenEfficiency.UI.YourDialog"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Your Dialog" Height="400" Width="300"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ShowInTaskbar="False" ResizeMode="NoResize"
        Topmost="True">

    <Border Background="#FFFFFF" BorderBrush="#CCCCCC" BorderThickness="1"
            CornerRadius="8" Padding="20">
        <!-- Your content here -->
    </Border>
</Window>
```

### 2. Create Code-Behind
```csharp
// UI/Dialogs/YourDialog.xaml.cs
using System.Windows;
using System.Windows.Interop;

namespace oPenEfficiency.UI
{
    public partial class YourDialog : Window
    {
        public YourDialog()
        {
            InitializeComponent();
        }

        // Optional: Enable drag for borderless windows
        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            var helper = new HwndSource(PresentationSource.FromVisual(this) as HwndSource);
            helper.AddHook(WndProc);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            // Handle WM_NCHITTEST for drag support
            return IntPtr.Zero;
        }
    }
}
```

### 3. Show the Dialog
```csharp
// In your feature or ThisAddIn.cs
var dialog = new YourDialog();

// Position near selection or cursor
var mousePos = System.Windows.Forms.Cursor.Position;
dialog.Left = mousePos.X;
dialog.Top = mousePos.Y;

// Enable owner for proper Z-order
dialog.Owner = System.Windows.Application.Current.MainWindow;
dialog.ShowDialog(); // Modal
// or dialog.Show(); // Modeless
```

---

## Adding a Floating Toolbar

Floating toolbars are used for quick actions on selected shapes (e.g., Harvey Balls, Star Ratings).

### Pattern: Create Inline
```csharp
private void ShowYourToolbar()
{
    var toolbar = new FloatingToolbar();

    // Add label (emoji or text)
    toolbar.AddLabel("\u2699\uFE0F"); // Gear emoji

    // Add button
    toolbar.AddButton("Action", () =>
    {
        // Handle click
        var manager = new PowerPointManager(Application);
        YourFeature.Execute(manager);
        toolbar.Close();
    });

    // Add separator
    toolbar.AddSeparator();

    // Add slider (0-100, initial value 50)
    toolbar.AddSlider(0, 100, 50, (s, e) =>
    {
        // Handle value change (live preview)
    });

    // Position near shape
    var window = Application.ActiveWindow;
    int pixelX = window.PointsToScreenPixelsX(shape.Left + shape.Width);
    int pixelY = window.PointsToScreenPixelsY(shape.Top);

    toolbar.Left = (pixelX * 96.0 / _dpiX) + 10;
    toolbar.Top = (pixelY * 96.0 / _dpiY);

    toolbar.Show();
}
```

### Pattern: Track Instance
For toolbars that should close when selection changes:
```csharp
// Field in ThisAddIn.cs
private YourToolbar _yourToolbar;

private void ShowYourToolbar(Shape shape)
{
    // Close existing toolbar
    if (_yourToolbar != null)
    {
        try { _yourToolbar.Close(); }
        catch (Exception ex) { ExceptionLogger.Log(ex, "CloseOldToolbar"); }
        _yourToolbar = null;
    }

    _yourToolbar = new YourToolbar();
    _yourToolbar.Closed += (s, args) => { _yourToolbar = null; };

    // Position and show
    _yourToolbar.Show();
}
```

---

## Modifying Existing UI

### Change Dialog Content
1. Open the `.xaml` file
2. Modify the XAML structure
3. Add corresponding event handlers in `.xaml.cs`

### Change Toolbar Items
1. Find the `Show*Toolbar` method in `ThisAddIn.cs`
2. Modify the `toolbar.Add*` calls

### Add New Option to Dialog
```xml
<!-- Add a new button -->
<Button Content="New Action" Click="NewAction_Click"/>
```

```csharp
// In code-behind
private void NewAction_Click(object sender, RoutedEventArgs e)
{
    var manager = new PowerPointManager(Application);
    YourFeature.Execute(manager);
    this.Close();
}
```

---

## Styling Conventions

### Colors
```csharp
// Dialog backgrounds (use Dark Theme resources)
Background: "{StaticResource WindowBackgroundBrush}"
Cards/Panels: "{StaticResource SurfaceBrush}"
Border: "{StaticResource BorderBrush}"

// Accent colors (use theme colors when possible)
Primary/Accent: "{StaticResource AccentBrush}"
Success: "#107C10" (Green)
Warning: "#FFB900" (Yellow)
Error: "#E81123" (Red)
```

### Buttons
- **Primary Actions:** `Style="{StaticResource AccentButton}"` (e.g., OK, Save)
- **Secondary Actions:** `Style="{StaticResource SecondaryButton}"` (e.g., Cancel, Close)

### Fonts
```csharp
// Default to system font (Segoe UI on Windows)
FontFamily: "Segoe UI"
FontSize: 14 (body), 12 (secondary), 16 (headers)
```

### Spacing
```csharp
// Consistent padding/margins
Control Padding: 20px
Item Margin: 8px vertical
Group Margin: 16px vertical
```

---

## Positioning Windows

### Near Shape
```csharp
int pixelX = window.PointsToScreenPixelsX(shape.Left + shape.Width);
int pixelY = window.PointsToScreenPixelsY(shape.Top);
dialog.Left = (pixelX * 96.0 / _dpiX) + 10;
dialog.Top = (pixelY * 96.0 / _dpiY);
```

### Near Cursor
```csharp
var mousePos = System.Windows.Forms.Cursor.Position;
dialog.Left = mousePos.X + 10;
dialog.Top = mousePos.Y + 10;
```

### Centered on Window
```csharp
var mainWindow = Application.Current.MainWindow;
dialog.Left = mainWindow.Left + (mainWindow.Width - dialog.Width) / 2;
dialog.Top = mainWindow.Top + (mainWindow.Height - dialog.Height) / 2;
```

---

## Handling Selection Changes

To close UI when selection changes:

```csharp
// In ThisAddIn.cs - ProcessSelectionChange method
private void ProcessSelectionChange(PowerPoint.Selection Sel)
{
    // Close toolbar if not relevant to new selection
    if (_yourToolbar != null)
    {
        // Check if new selection is still relevant
        bool stillRelevant = CheckIfRelevant(Sel);

        if (!stillRelevant)
        {
            try { _yourToolbar.Close(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "CloseToolbar"); }
            _yourToolbar = null;
        }
    }
}
```

---

## DPI Handling

Always account for system DPI when positioning:

```csharp
// Capture DPI at startup (in ThisAddIn_Startup)
private float _dpiX = 96f, _dpiY = 96f;

using (System.Drawing.Graphics g = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
{
    _dpiX = g.DpiX;
    _dpiY = g.DpiY;
}

// Use in positioning calculations
double wpfUnits = (screenPixels * 96.0) / dpi;
```

---

## Common UI Patterns

### Confirmation Dialog
```csharp
var result = System.Windows.Forms.MessageBox.Show(
    "Are you sure you want to proceed?",
    "Confirm Action",
    MessageBoxButtons.YesNo,
    MessageBoxIcon.Question);

if (result == System.Windows.Forms.DialogResult.Yes)
{
    // Proceed
}
```

### Error Message
```csharp
System.Windows.Forms.MessageBox.Show(
    $"Operation failed: {ex.Message}",
    "Error",
    MessageBoxButtons.OK,
    MessageBoxIcon.Error);
```

### Progress Indicator
For long operations, show a modeless progress window:
```csharp
var progress = new ProgressWindow("Processing...");
progress.Show();

Task.Run(() =>
{
    // Long operation
    progress.Dispatcher.Invoke(() => progress.Close());
});
```

---

## Testing UI Changes

1. **Visual Verification:** Run the add-in and trigger the UI
2. **DPI Testing:** Test at 100%, 125%, 150% system DPI
3. **Selection Changes:** Verify UI closes when expected
4. **Keyboard Navigation:** Ensure Tab and Enter work correctly
5. **Accessibility:** Verify screen readers can read controls

---

## Troubleshooting

### Window Doesn't Show
- Check if WPF Application is initialized (`EnsureWpfApplication()`)
- Verify `Show()` or `ShowDialog()` is called
- Check `Owner` property for proper Z-order

### Wrong Position
- Verify DPI calculation: `(pixels * 96) / dpi`
- Use `PointsToScreenPixelsX/Y` for PowerPoint coordinates
- Check if window is positioned off-screen

### UI Doesn't Close
- Ensure `Closed +=` handler is properly attached
- Check for circular references preventing GC
- Call `Close()` explicitly

### Blank/Empty Controls
- Verify data binding context
- Check if items are added before `Show()`
- Ensure dispatcher thread affinity for WPF operations

---

## Selection-Based Icon Disabling (Added 2026-03-13)

Feature buttons in the sidebar are automatically enabled/disabled based on the current selection.

### How It Works

1. **FeatureDisplayInfo** (in `UI/FeatureLibrary.cs`) defines selection requirements:
   ```csharp
   public struct FeatureDisplayInfo
   {
       public int MinSelection { get; set; }       // Minimum shapes needed
       public int MaxSelection { get; set; }       // Maximum shapes (0 = unlimited)
       public PowerPoint.PpSelectionType RequiredType { get; set; }
       public bool RequiresTable { get; set; }     // Requires a table selection
   }
   ```

2. **MainSidebar** tracks selection and updates button states:
   ```csharp
   private PowerPoint.Selection _currentSelection;

   public void OnSelectionChange(PowerPoint.Selection sel)
   {
       _currentSelection = sel;
       UpdateButtonStates();  // Updates all button enabled states
   }
   ```

3. **IsFeatureEnabled** checks requirements against selection:
   - Returns `false` if selection is null
   - Checks `RequiredType` (e.g., must be shape selection)
   - Checks `RequiresTable` (iterates ShapeRange for msoTable)
   - Checks `MinSelection`/`MaxSelection` counts

4. **Visual feedback**: Disabled buttons show 40% opacity and arrow cursor (defined in `MainSidebar.xaml` GridIconButton style)

### Adding New Features

When defining a feature in `FeatureLibrary.GetFeatureInfo()`:

```csharp
// Alignment features (need 2+ shapes)
case "BtnAlignLeft":
    return new FeatureDisplayInfo {
        ...,
        MinSelection = 2,
        RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes
    };
```

**IMPORTANT:** When adding a new feature, you MUST also add its entry to the `FeatureLibrary.AllFeatures` static list in **alphabetical order by Name**. This ensures the "Available Features" pool in the Layout settings tab remains organized and searchable for the user.

// Table features
case "BtnTableSort":
    return new FeatureDisplayInfo {
        ...,
        RequiresTable = true
    };

// Single-shape features
case "BtnHarveyBall":
    return new FeatureDisplayInfo {
        ...,
        MinSelection = 1,
        MaxSelection = 1,
        RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes
    };
```

### Important Notes

- Features without explicit requirements will be enabled for any selection
- Table detection iterates `selection.ShapeRange[i].Type == MsoShapeType.msoTable`
- The `IsEnabled` trigger in XAML must be present for visual feedback

---

## Context Menus for Feature Buttons

Feature buttons in the sidebar can have right-click context menus for additional options.

### Architecture

Context menus are attached in `MainSidebar.xaml.cs` using two method overloads:

```csharp
private void AttachStandardContextMenu(Button btn)
private void AttachStandardContextMenu(ToggleButton btn)
```

### Critical Distinction: Button vs ToggleButton

**Regular Button (99% of features):**
- Created via `CreateFeatureButton()` in `MainSidebar.xaml.cs`
- Used for standard click-to-execute features
- Context menus go in `AttachStandardContextMenu(Button btn)`

**ToggleButton (only 2 features):**
- Created via `CreateToggleButton()` in `MainSidebar.xaml.cs`
- Used for toggleable features with on/off state
- Only `BtnSnapToObjects` and `BtnSnapToGrid` use ToggleButton
- Context menus go in `AttachStandardContextMenu(ToggleButton btn)`

### How to Add a Context Menu

1. **Determine button type:** Check `FeatureLibrary.GetFeatureInfo()` - if `IsToggle = true`, it's a ToggleButton

2. **Add to correct overload:**
   - Regular features → `AttachStandardContextMenu(Button btn)`
   - Toggle features → `AttachStandardContextMenu(ToggleButton btn)` (rarely needed)

3. **Add menu items:**

```csharp
private void AttachStandardContextMenu(Button btn)
{
    btn.ContextMenu = new ContextMenu();

    // Always add "Configure Shortcuts"
    var item = new MenuItem { Header = "Configure Shortcuts" };
    item.Click += (s, e) => {
        var win = new ShortcutConfigWindow(btn.Name);
        win.ShowDialog();
    };
    btn.ContextMenu.Items.Add(item);

    // Add feature-specific options
    if (btn.Name == "BtnYourFeature")
    {
        btn.ContextMenu.Items.Add(new Separator());

        var mi = new MenuItem { Header = "Option Name" };
        mi.Click += (s, e) => YourFeature.Execute(GetManager());
        btn.ContextMenu.Items.Add(mi);
    }
}
```

### Common Mistake to Avoid

**DO NOT add context menus for regular Button features in the ToggleButton overload.** This was a historical bug - the ToggleButton method contained context menus for 100+ Button features, which meant those menus never appeared.

**Rule of thumb:** If you're adding a context menu for any feature other than `BtnSnapToObjects` or `BtnSnapToGrid`, it goes in the `Button` overload.

---

## Opening Native PowerPoint Dialogs

When opening native PowerPoint dialogs/panes programmatically in VSTO add-ins, be aware of significant API limitations.

### Limitations

PowerPoint's VSTO interop API **does not reliably support** opening all native dialogs:

| Dialog/Pane | Programmatic Support | Notes |
|-------------|---------------------|-------|
| Format Shape | ❌ Unreliable | `ExecuteMso` often fails |
| Paragraph Dialog | ⚠️ Limited | Works in some versions |
| Font Dialog | ⚠️ Limited | Depends on context |

The root causes:
1. **CommandBars.ExecuteMso()** - Many command IDs are undocumented or version-specific
2. **Win32MouseHook** - Our global mouse hook (used for Snap-to-Objects/Grid) intercepts `SendKeys.SendWait()` calls, making keyboard shortcut fallbacks unreliable

### Fallback Pattern for Format Shape

```csharp
case "BtnFormatShape":
    try
    {
        // Check if a shape is selected
        var shapes = manager.GetSelectedShapes();
        if (shapes == null || shapes.Count == 0)
        {
            MessageBox.Show("Please select a shape first.");
            break;
        }

        var app = manager.GetApplication();
        var commandBars = app.CommandBars;

        // Method 1: Try ExecuteMso with documented command ID
        try
        {
            commandBars.ExecuteMso("FormatObjectDialog");
            return; // Success
        }
        catch (COMException ex1)
        {
            Debug.WriteLine($"FormatObjectDialog failed: {ex1.Message}");

            // Method 2: Search for Format Shape control by name
            foreach (CommandBar bar in commandBars)
            {
                foreach (CommandBarControl control in bar.Controls)
                {
                    string ctrlName = control.Caption?.ToLower() ?? "";
                    if (ctrlName.Contains("format") && ctrlName.Contains("shape"))
                    {
                        control.Execute();
                        return;
                    }
                }
            }
        }
    }
    catch (Exception ex)
    {
        Debug.WriteLine($"FormatShape failed: {ex.Message}");
    }

    // Fallback: Inform user how to open manually
    MessageBox.Show(
        "Cannot open Format Shape pane automatically.\n\n" +
        "Please use: Ctrl+1 or Right-click → Format Shape",
        "Format Shape", MessageBoxButtons.OK, MessageBoxIcon.Information);
    break;
```

### Known Working ExecuteMso Commands

These command IDs are confirmed working in PowerPoint 2010+:

| Command ID | Action |
|------------|--------|
| `"Paste"` | Standard paste |
| `"Cut"` | Standard cut |
| `"Copy"` | Standard copy |
| `"Undo"` | Undo last action |
| `"Redo"` | Redo last action |

### Commands That Often Fail

| Command ID | Issue |
|------------|-------|
| `"FormatObjectDialog"` | Returns "value out of range" error |
| `"FormatShape"` | Command not found |
| `"ParagraphDialog"` | Version-dependent |

### Alternative Approaches

If `ExecuteMso` fails:

1. **SendKeys** - May be intercepted by global hooks (like `Win32MouseHook`)
2. **CommandBar control search** - Iterate controls by name, execute directly
3. **Custom WPF UI** - Build your own dialog/panel with the needed controls
4. **User instructions** - Guide users to the native shortcut (e.g., Ctrl+1)

### Recommendation

For features requiring extensive shape formatting, consider building a custom WPF panel that uses PowerPoint interop directly (e.g., `shape.Fill.ForeColor`, `shape.Line.Weight`) rather than relying on native dialogs.

---

## Locking Guideline Shapes

When creating guideline line shapes (like in SlideGuidelinesFeature), always lock them to prevent accidental manipulation.

### LockGuideline Pattern

```csharp
/// <summary>
/// Applies locking properties to a guideline shape to prevent accidental manipulation.
/// Note: PowerPoint doesn't have a true "lock position" property for shapes via interop.
/// We use available locking options and mark the shape for detection in selection handlers.
/// </summary>
private static void LockGuideline(Shape guideline)
{
    try
    {
        // Lock aspect ratio (prevents proportional resizing)
        guideline.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

        // Mark as guideline in AlternativeText for detection in selection handlers
        guideline.AlternativeText = "oPE_Guideline";

        // Add tag for detection (line shapes may not support Tags in all PowerPoint versions)
        try
        {
            guideline.Tags.Add("oPE_Locked", "true");
        }
        catch
        {
            // Line shapes may not support Tags in some versions - AlternativeText is still set
        }
    }
    catch (Exception ex)
    {
        ExceptionLogger.Log(ex, "LockGuideline");
    }
}
```

### Usage in Guide Creation

```csharp
var line = slide.Shapes.AddLine(positionPts, 0, positionPts, (float)presentation.PageSetup.SlideHeight);
line.Name = $"Guide_V_{name}";
line.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(100, System.Drawing.Color.Blue));
line.Line.Weight = 1f;
line.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;

// Lock the shape to prevent accidental manipulation
LockGuideline(line);

line.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);
```

### Important Notes

- PowerPoint's interop API does NOT expose a true "lock position" property for shapes
- `LockAspectRatio` only prevents proportional resizing, not movement
- Use `AlternativeText` and `Tags` to mark shapes for detection in `ProcessSelectionChange` handlers
- Consider adding logic to `ThisAddIn.cs` to auto-deselect or warn when guidelines are selected
