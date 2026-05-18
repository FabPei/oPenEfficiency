# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

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

## Architecture Overview

**Entry Point:** `ThisAddIn.cs` - VSTO add-in lifecycle, event handlers, ribbon callbacks

**Core Pattern:** All PowerPoint COM operations flow through `PowerPointManager`:
```csharp
var manager = new PowerPointManager(this.Application);
var shapes = manager.GetSelectedShapes();
```

**Project Structure:**
```
oPenEfficiency/
├── ThisAddIn.cs              # VSTO entry point, selection events, shortcut registration
├── UI/
│   ├── MainRibbon.cs/.xml    # Custom ribbon definition
│   ├── Panels/               # Docked sidebar (MainSidebar.xaml)
│   ├── Dialogs/              # Modal WPF dialogs
│   ├── Toolbars/             # Floating toolbars (FloatingToolbar.cs)
│   └── Styles/               # Controls.xaml, Theme.xaml (dark theme)
├── Features/                 # Static feature classes by category
│   ├── Alignment/            # Align, Match Size, Swap Positions
│   ├── Tables/               # Convert, Transpose, Sort, Split
│   ├── Text/                 # Split by Paragraphs, Replace Font
│   ├── Visuals/              # Pick Color, Remove Gaps, Object Connector
│   └── Utilities/            # AgendaWizard, StickyNote, ShapeLocking
├── Services/
│   ├── PowerPointManager.cs  # Central PowerPoint access
│   ├── ShortcutManager.cs    # Global hotkey management
│   ├── AgendaService.cs      # CustomXML persistence
│   └── StyleScanner.cs       # Presentation style analysis
└── Utils/
    ├── ExceptionLogger.cs    # Logs to %TEMP%\oPenEfficiency\
    └── Win32MouseHook.cs     # Global mouse hook for drag operations
```

## Feature Implementation Pattern

**New Pattern (Auto-Discovered):** Features with `[FeatureMetadata]` attribute are auto-discovered:
```csharp
namespace oPenEfficiency.Features.Alignment
{
    [FeatureMetadata(
        Id = "BtnAlignLeft",
        Name = "Align Left",
        Tooltip = "Align left edges",
        IconData = "M4,2H2V22H4V2M22,10H6V14H22V10Z",
        Color = "#10B981",
        Description = "Aligns objects to left edge",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class AlignToFirstLeftFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            var shapes = manager.GetSelectedShapes();
            // Implementation
        }
    }
}
```

**Benefits:** No manual registration needed. The `FeatureDiscovery` service auto-finds all attributed classes via reflection.

**Legacy Pattern (Manual Switch):** Features without attributes still work via manual switch in `MainSidebar.xaml.cs`.

**Adding a new feature (auto-discovered):**
1. Create class in `Features/` folder with `[FeatureMetadata]` attribute
2. Add button ID to `UI/MainRibbon.xml` (if using ribbon)
3. Add to `SidebarFeature.AllFeatures` list (for sidebar display)
4. Done! No need to edit `GetFeatureInfo` or `ExecuteFeature` switches

**Adding a new feature (legacy, no attribute):**
1. Create class in `Features/` folder
2. Add to `MainSidebar.xaml.cs` manual `ExecuteFeature` switch
3. Add to `FeatureLibrary.GetFeatureInfo` switch for metadata
4. Add button to sidebar/ribbon

## Selection-Based Feature Enabling

Features define requirements in `FeatureDisplayInfo`:
- `MinSelection`/`MaxSelection` - shape count constraints
- `RequiredType` - e.g., `ppSelectionShapes`
- `RequiresTable` - detects table selection

`MainSidebar.OnSelectionChange()` updates button enabled states automatically.

## WPF UI Patterns

**Dark Theme:** All dialogs use `Controls.xaml` and `Theme.xaml` resource dictionaries.

**Floating Windows:** Follow the pattern in `.agents/skills/floating-window-creation.md`:
- `WindowStyle="None"`, `AllowsTransparency="True"`
- Root `Border` with `CornerRadius="12"` and `DropShadowEffect`
- `Window_MouseDown` for drag-to-move
- Use `StaticResource` brushes (never hardcode colors)

**Dialogs:** Inherit from `Window`, positioned near cursor or shape:
```csharp
var mousePos = System.Windows.Forms.Cursor.Position;
dialog.Left = mousePos.X;
dialog.Top = mousePos.Y;
dialog.ShowDialog();
```

**Floating Toolbars:** Use `FloatingToolbar` helper:
```csharp
var toolbar = new FloatingToolbar();
toolbar.AddLabel("\u2699\uFE0F");
toolbar.AddButton("Action", () => { /* handler */ });
toolbar.Show();
```

## Critical Implementation Details

**WPF Application Singleton:** Created in `ThisAddIn_Startup`:
```csharp
if (System.Windows.Application.Current == null)
{
    var app = new System.Windows.Application();
    app.ShutdownMode = ShutdownMode.OnExplicitShutdown;
}
```

**DPI-Aware Positioning:**
```csharp
// Capture at startup
using (var g = Graphics.FromHwnd(IntPtr.Zero))
{
    _dpiX = g.DpiX;
    _dpiY = g.DpiY;
}

// Convert pixels to WPF units
toolbar.Left = (pixelX * 96.0 / _dpiX);
```

**Shape Metadata:** Use `Tags` collection with `oPE_` prefix:
```csharp
```

**State Persistence:** Store in `AlternativeText`:
```csharp
group.AlternativeText = $"{HarveyBallTag}|{percentage}";
```

**Exception Handling:** Always log:
```csharp
catch (Exception ex)
{
    ExceptionLogger.Log(ex, "FeatureName.Execute");
}
```

## CI/CD

**Build Workflow:** `.github/workflows/build.yml` - triggers on push to main/master/Refactor

**Publish Workflow:** `.github/workflows/publish.yml` - creates GitHub Release when `ApplicationVersion` changes in `.csproj`

**Release Process:**
1. Update `<ApplicationVersion>` in `oPenEfficiency.csproj`
2. Commit and push - workflow creates tag and release with `oPenEfficiency-Installer.zip`

## Testing

- **Manual testing:** Required - PowerPoint VSTO add-ins cannot be unit tested easily
- **Debug logs:** `%TEMP%\oPenEfficiency\error_YYYYMMDD.log`
- **Shortcut debug:** `%TEMP%\oPenEfficiency\ShortcutDebug_oPenEfficiency.txt`

## Key Files

| File | Purpose |
|------|---------|
| `ThisAddIn.cs` | Entry point, event handling, shortcut registration |
| `UI/MainRibbon.xml` | Ribbon XML definition |
| `UI/FeatureLibrary.cs` | Feature metadata and selection requirements |
| `UI/Panels/MainSidebar.xaml.cs` | Sidebar UI, feature execution, selection-based enabling |
| `UI/Styles/Controls.xaml` | Centralized WPF styles (dark theme) |
| `Services/PowerPointManager.cs` | PowerPoint COM wrapper |
| `Services/FeatureDiscovery.cs` | Auto-discovers features with `[FeatureMetadata]` attribute |
| `Services/FeatureWrapper.cs` | Executes static feature classes via reflection |
| `Services/Attributes/FeatureMetadataAttribute.cs` | Attribute definition for feature metadata |
| `Utils/ExceptionLogger.cs` | Centralized error logging |

## Maintainer's To-Do List (Open-Source Improvements)

**Priority 1 - Quick Wins (COMPLETED):**

- [x] **Create `.editorconfig`** - Enforce coding standards automatically
- [x] **Create feature template** - `.templates/NewFeature.cs` with copy-paste structure
- [x] **Fix AttachContextMenu** - Merged Button/ToggleButton overloads into single method

**Priority 2 - Documentation (COMPLETED):**

- [x] **Feature browser table** - Added to README with all 90+ features (Button ID, selection, dialog)
- [x] **XML comments on weird patterns** - Added to CLAUDE.md and source files:
  - `oPE_` tag prefix - Avoids conflicts, identifies oPenEfficiency shapes
  - `AlternativeText` for state - Only property surviving copy/paste
  - Win32MouseHook intercepts SendKeys - Why Format Shape dialog fallback fails
  - 1-indexed COM collections - PowerPoint quirk documented
  - `ButtonBase` for context menus - Prevents overload bugs

**Priority 3 - Architecture (COMPLETED):**

- [x] **Attribute-based registration** - Auto-discovery via `[FeatureMetadata]` attribute
  - Created `Services/FeatureMetadataAttribute.cs` - Attribute definition
  - Created `Services/FeatureWrapper.cs` - Reflection-based execution wrapper
  - Created `Services/FeatureDiscovery.cs` - Assembly scanner for auto-discovery
  - Updated `FeatureLibrary.GetFeatureInfo()` - Auto-discovery with manual fallback
  - Updated `MainSidebar.ExecuteFeature()` - Executes auto-discovered features first
  - Updated `oPenEfficiency.csproj` - Glob pattern `<Compile Include="Features\**\*.cs" />`
  - Migrated `AlignToFirstLeftFeature` as migration example

**What this fixes:**
- No more forgotten registrations - features with attributes are auto-registered
- Reduces boilerplate from editing 4 files to 1 (just the feature class)
- `GetFeatureInfo()` and `ExecuteFeature()` work for both attributed and non-migrated features
- New feature files auto-included via glob pattern in csproj

---

## Migration Plan: Full Attribute-Based Auto-Discovery

**Status:** **COMPLETED** - All 8 batches migrated (~90 features).

**Result:** Adding new features now requires only 1 file (feature class + attribute).

### Migration Approach

**Batch by category** (easiest to hardest):

#### Batch 1: Simple Alignment Features (~17 features) **COMPLETED**
All migrated: AlignToFirstLeftFeature, AlignToFirstRightFeature, AlignToFirstTopFeature, AlignToFirstBottomFeature, AlignToFirstCenterHorizontalFeature, AlignToFirstCenterVerticalFeature, AlignRightFeature, MatchSizeFeature, MatchWidthToFirstFeature, MatchHeightToFirstFeature, SwapPositionsFeature, SwapPositionsHorizontalFeature, SwapPositionsVerticalFeature, DockToFirstLeftFeature, DockToFirstRightFeature, DockToFirstTopFeature, DockToFirstBottomFeature
- AlignShapesFeature - has string parameter (wrapper added)

#### Batch 2: Stretch Features (~8 features) **COMPLETED**
All migrated: StretchToFirstLeftFeature, StretchToFirstRightFeature, StretchToFirstTopFeature, StretchToFirstBottomFeature, StretchToFirstLeftEdgeFeature, StretchToFirstRightEdgeFeature, StretchToFirstTopEdgeFeature, StretchToFirstBottomEdgeFeature

#### Batch 3: Visuals (~16 features) **COMPLETED**
All migrated: PickColorFeature, RectifyRotationFeature, ThemeColorPickerFeature, TransparentColorFeature, RemoveHorizontalGapsFeature, RemoveVerticalGapsFeature, AdjustHorizontalSpacingFeature (wrapper), AdjustVerticalSpacingFeature (wrapper), SplitShapeToGridFeature (wrapper), RepeatShapeFeature (wrapper), AdvancedSelectionFeature (wrapper), AddStickerFeature, IllustrativeStickerFeature (wrapper), ObjectConnectorFeature (wrapper), CheckboxFeature (wrapper)
- SyncObjectsFeature - migrated (manual dialog handling)
- StyleCheckFeature - migrated (manual dialog handling)

#### Batch 4: Utilities (~25 features) **COMPLETED**
All migrated: HideSelectedFeature, ShowHiddenFeature, SelectSameTypeFeature, CopyCoordinatesFeature, PasteCoordinatesFeature, AnonymizeFeature, DimensionLockFeature (wrapper), QRCodeFeature (wrapper), PropertyExtractionFeature (wrapper), ShapeLockingFeature (wrapper), StickyNoteFeature, StickyNoteManagerFeature (wrapper), LotteryFeature (manual dialog), WinnerPickerFeature (manual dialog), NumerationFeature (wrapper), CleanerFeature (manual dialog), ActionTitlesFeature (manual dialog), SlideGuidelinesFeature (wrapper), UpdateExcelChartsFeature (wrapper), AgendaWizardFeature (manual dialog), SaveAgendaLayoutFeature (manual dialog)
- VisibilityFeature - Helper class (no migration needed)
- SnapToObjectsFeature - migrated (toggle, manual handling)
- SnapToGridFeature - migrated (toggle, manual handling)
- FlightModeFeature - migrated (toggle, manual handling)

#### Batch 5: Text (7 features) **COMPLETED**
All migrated: DeleteTextFeature, InsertTextAtCursorFeature (wrapper), SwapTextPlainFeature, SwapTextFormattedFeature, SplitByParagraphsFeature, FormatNumbersFeature (manual dialog), ReplaceFontFeature (manual dialog)

#### Batch 6: Charts (8 features) **COMPLETED**

#### Batch 7: Tables (17 features + 1 helper) **COMPLETED**
All migrated: ConvertShapesToTableFeature, ConvertTableToShapesFeature, TableSortFeature (wrapper), TableFormatPainterFeature (manual dialog), TableColumnWidthFeature (wrapper), TableRowHeightFeature (wrapper), TableDimensionsFeature, TableTransposeFeature, TableSplitFeature (wrapper), TableSumFeature (wrapper), InsertTableColumnLeftFeature (wrapper), InsertTableColumnRightFeature (wrapper), InsertTableRowTopFeature (wrapper), InsertTableRowBottomFeature (wrapper), TableColumnInsertionFeature (wrapper), TableRowInsertionFeature (wrapper), TableBrandingFeature (manual dialog)
- TableHelper.cs - Helper class (no migration needed)

#### Batch 8: Special/Arrange Features (6 features) **COMPLETED**
All migrated: ArrangeProFeature (manual dialog), ArrangeGridFeature (wrapper), ArrangeInShapeFeature (wrapper), AlignToTableCellFeature (manual dialog)
- SmartCornersFeature - needs migration
- ShapeLibraryManager, SlideLibraryManager, ImageLibraryManager - Helper classes (no migration needed)
### Migration Pattern

For each feature file:

1. **Add using statement:**
   ```csharp
   using oPenEfficiency.Services.Attributes;
   ```

2. **Add attribute above class:**
   ```csharp
   [FeatureMetadata(
       Id = "BtnFeatureName",
       Name = "Display Name",
       Tooltip = "Short tooltip",
       IconData = "M12,2L14,4...",
       Color = "#10B981",
       Description = "Full description",
       MinSelection = 1,
       RequiredType = PpSelectionType.ppSelectionShapes)]
   public static class FeatureName { ... }
   ```

3. **Lookup metadata** from existing `FeatureLibrary.GetFeatureInfo()` switch (copy values).

4. **Build and test** after each batch.

### Special Cases

**Features with parameters** (e.g., `Execute(manager, "left")`):
- Option A: Keep manual switch in MainSidebar.xaml.cs
- Option B: Create overload `Execute(PowerPointManager)` that calls with default params

**Toggle features** (SnapToObjects, SnapToGrid, FlightMode):
- Need special handling - Execute takes ToggleButton parameter
- Recommend keeping manual switch execution

**Dialog features** (opens window/dialog):
- Can migrate attribute for metadata
- Keep manual switch for execution (or create simple Execute wrapper)

**Features returning void vs bool**:
- Auto-discovery expects `bool Execute(PowerPointManager)`
- Most features already return bool

### Rollback Plan

If migration causes issues:
1. Comment out `FeatureDiscovery.Initialize()` call in `FeatureLibrary.cs`
2. Existing manual switch statements remain functional
3. Migration is additive - no breaking changes to existing features

### Progress Tracking

| Batch | Features | Status |
|-------|----------|--------|
| Batch 1 | Alignment (17) | **COMPLETED** (17/17) |
| Batch 2 | Stretch (8) | **COMPLETED** (8/8) |
| Batch 3 | Visuals (16) | **COMPLETED** (14/16 - 2 manual for dialogs) |
| Batch 4 | Utilities (25+) | **COMPLETED** (~16/25 migrated, ~9 manual/toggles) |
| Batch 5 | Text (7) | **COMPLETED** (7/7) |
| Batch 6 | Charts (8) | **COMPLETED** (8/8 - 1 manual for dialog) |
| Batch 7 | Tables (17) | **COMPLETED** (17/17 - 1 helper) |
| Batch 8 | Special (18) | **COMPLETED** (18/18 - 5 manual dialogs, 3 toggles, 1 helper) |

**Total:** ~110 features migrated (~90 with attributes, ~18 manual/toggles, ~8 helpers)

### Completed Features

**Batch 1 - Alignment (17 features):**
- AlignToFirstLeftFeature, AlignToFirstRightFeature, AlignToFirstTopFeature, AlignToFirstBottomFeature
- AlignToFirstCenterHorizontalFeature, AlignToFirstCenterVerticalFeature
- AlignRightFeature, AlignShapesFeature (manual - has parameter)
- MatchSizeFeature, MatchWidthToFirstFeature, MatchHeightToFirstFeature
- SwapPositionsFeature, SwapPositionsHorizontalFeature, SwapPositionsVerticalFeature
- DockToFirstLeftFeature, DockToFirstRightFeature, DockToFirstTopFeature, DockToFirstBottomFeature

**Batch 2 - Stretch (8 features):**
- StretchToFirstLeftFeature, StretchToFirstRightFeature, StretchToFirstTopFeature, StretchToFirstBottomFeature
- StretchToFirstLeftEdgeFeature, StretchToFirstRightEdgeFeature, StretchToFirstTopEdgeFeature, StretchToFirstBottomEdgeFeature

**Batch 3 - Visuals (14 features migrated, 2 manual):**
- PickColorFeature, RectifyRotationFeature, ThemeColorPickerFeature, TransparentColorFeature
- RemoveHorizontalGapsFeature, RemoveVerticalGapsFeature
- AdjustHorizontalSpacingFeature, AdjustVerticalSpacingFeature (wrappers)
- SplitShapeToGridFeature (wrapper), RepeatShapeFeature (wrapper)
- AdvancedSelectionFeature (wrapper), AddStickerFeature, IllustrativeStickerFeature (wrapper)
- ObjectConnectorFeature (wrapper), CheckboxFeature (wrapper)
- *Manual (dialog): StyleCheckFeature, SyncObjectsFeature*

**Batch 4 - Utilities (16 features migrated):**
- HideSelectedFeature, ShowHiddenFeature, SelectSameTypeFeature
- CopyCoordinatesFeature, PasteCoordinatesFeature
- AnonymizeFeature, DimensionLockFeature (wrapper), QRCodeFeature (wrapper)
- PropertyExtractionFeature (wrapper), ShapeLockingFeature (wrapper), StickyNoteFeature

**Batch 5 - Text (7 features):**
- DeleteTextFeature, InsertTextAtCursorFeature (wrapper), SwapTextPlainFeature, SwapTextFormattedFeature
- FormatNumbersFeature (dialog), ReplaceFontFeature (dialog), SplitByParagraphsFeature

**Batch 6 - Charts (8 features):**
- HarveyBallFeature (wrapper), TrafficLightFeature (wrapper), ThermometerFeature (wrapper)
- StarRatingFeature (wrapper), ProgressSeriesFeature (wrapper)

**Batch 7 - Tables (17 features migrated, 1 helper):**
- ConvertShapesToTableFeature, ConvertTableToShapesFeature
- TableSortFeature (wrapper), TableFormatPainterFeature (dialog)
- TableColumnWidthFeature (wrapper), TableRowHeightFeature (wrapper)
- TableDimensionsFeature, TableTransposeFeature, TableSplitFeature (wrapper)
- TableSumFeature (wrapper)
- InsertTableColumnLeftFeature (wrapper), InsertTableColumnRightFeature (wrapper)
- InsertTableRowTopFeature (wrapper), InsertTableRowBottomFeature (wrapper)
- TableColumnInsertionFeature (wrapper), TableRowInsertionFeature (wrapper)
- TableBrandingFeature (manual - opens dialog)
- *Helper class (no migration): TableHelper.cs*

**Batch 8 - Special/Arrange (5 features migrated):**
- ArrangeProFeature (manual dialog), ArrangeGridFeature (wrapper), ArrangeInShapeFeature (wrapper)
- AlignToTableCellFeature (manual dialog), SmartCornersFeature (wrapper)
- SyncObjectsFeature (manual dialog), StyleCheckFeature (manual dialog)
- StickyNoteManagerFeature (wrapper), ActionTitlesFeature (manual dialog)
- CleanerFeature (manual dialog), SlideGuidelinesFeature (wrapper)
- UpdateExcelChartsFeature (wrapper), NumerationFeature (wrapper)
- AgendaWizardFeature (manual dialog), SaveAgendaLayoutFeature (manual dialog)
- LotteryFeature (manual dialog), WinnerPickerFeature (manual dialog)
- SnapToGridFeature (toggle - manual), SnapToObjectsFeature (toggle - manual), FlightModeFeature (toggle - manual)
- *Helper classes (no migration): VisibilityFeature, ShapeLibraryManager, SlideLibraryManager, ImageLibraryManager*

**What NOT to do (solves problems we don't have):**

- [ ] `IFeature` interface - Over-engineering for a VSTO add-in with no plugin plans
- [ ] Unit test framework - VSTO is fundamentally untestable; manual testing is fine
- [ ] Dependency injection - Overkill for 100 static classes
- [ ] Microservice architecture - It's a desktop add-in

---

## Additional:

Always ask Questions if the Input from the user is unclear or ambigous. Always answer with "Okay chef".
