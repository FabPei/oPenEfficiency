# openEfficiency feature reference

> **Source reference.** This page documents the PowerPoint commands and widgets in this source checkout. Published releases and a newer source checkout can differ; match the release tag and notes to your installed version. The current MSI release is [v26.10.23](https://github.com/FabPei/oPenEfficiency/releases/tag/v26.10.23).

Use the [release notes](https://github.com/FabPei/oPenEfficiency/releases) for the features available in your installed build. Find commands through Feature Explorer, Command Search or Settings; catalog-only tools may need to be added to your sidebar. The three widgets are embedded sidebar controls. Labs tools are experimental. Selection and context-menu options depend on the command.

OneNote is outside this PowerPoint reference and remains paused. Existing Gantt/Timeline commands are documented as source behavior; additional Gantt work and its outstanding verification are deferred to a later release. Documentation coverage is not evidence that a feature has passed runtime or visual checks.

## Browse by task

[Alignment and size](#alignment-and-size) · [Arrange and organize](#arrange-and-organize) · [Text and formatting](#text-and-formatting) · [Tables](#tables) · [Visual elements](#visual-elements) · [Storyline and review](#storyline-and-review) · [Data and automation](#data-and-automation) · [Libraries](#libraries) · [Settings and help](#settings-and-help) · [Labs](#labs)

## Alignment and size

### Adjust Spacing

**Command ID:** `BtnAdjustSpacing` · **Type:** Command · **Maturity:** Stable

Opens anchored and center-preserving horizontal and vertical spacing actions for the selected shapes.

**How to use**

Select at least two shapes, click the feature, then choose whether to increase or decrease horizontal or vertical spacing. The centered decrease actions move shapes equally toward the selection midpoint instead of anchoring the first spatial shape. The configured spacing increment is applied between each neighboring shape. A keyboard shortcut assigned to this command directly increases horizontal spacing.

### Align Bottom

**Command ID:** `BtnAlignBottom` · **Type:** Command · **Maturity:** Stable

Multi-select: aligns bottom edges to the anchor object. Single select: snaps bottom edge to the nearest horizontal guide line (or slide edge). Right-click for 'Move Until Object'.

**How to use**

With multiple shapes, aligns their bottom edges to the first or last selected anchor configured in Settings. With one shape, snaps its bottom edge to the nearest lower horizontal guide or to the slide bottom. Right-click for explicit selection, slide, first-anchor, last-anchor, and collision modes. Each action creates a PowerPoint Undo entry.

### Align Center

**Command ID:** `BtnAlignCenter` · **Type:** Command · **Maturity:** Stable

Multi-select: centers shapes horizontally on the configured first/last anchor. Single select: snaps to a vertical guide or slide center.

**How to use**

With multiple selected shapes, keeps the anchor object configured in Settings (**First selected** or **Last selected**) fixed and moves every other shape so their horizontal centers match. With one shape, snaps its horizontal center to the nearest vertical presentation, master, or layout guide; the slide center is the fallback. Right-click to explicitly align within the selection, center every shape on the slide, snap each shape to its nearest guide, or use the configured anchor. Each action creates a PowerPoint Undo entry.

### Align Center Vertical

**Command ID:** `BtnAlignCenterVertical` · **Type:** Command · **Maturity:** Stable

Multi-select: centers shapes vertically on the configured first/last anchor. Single select: snaps to a horizontal guide or slide middle.

**How to use**

With multiple selected shapes, keeps the anchor object configured in Settings (**First selected** or **Last selected**) fixed and moves every other shape so their vertical centers match. With one shape, snaps its vertical center to the nearest horizontal presentation, master, or layout guide; the slide middle is the fallback. Right-click to explicitly align within the selection, center every shape on the slide, snap each shape to its nearest guide, or use the configured anchor. Each action creates a PowerPoint Undo entry.

### Align Diagonally

**Command ID:** `BtnDiagonalAlign` · **Type:** Command · **Maturity:** Stable

Aligns selected shapes in a diagonal line (e.g., 45 degrees). Use the right-click menu to choose specific angles.

**How to use**

Places the selected shapes in an evenly spaced line at the chosen angle. The toolbar lets you use common angles and choose whether the first shape, last shape, or the selection center remains anchored.

### Align Left

**Command ID:** `BtnAlignLeft` · **Type:** Command · **Maturity:** Stable

Multi-select: aligns left edges to the configured first/last anchor. Single select: snaps to the nearest left-side vertical guide or slide edge.

**How to use**

With multiple selected shapes, keeps the anchor object configured in Settings (**First selected** or **Last selected**) fixed and aligns every other left edge to it. With one shape, snaps the left edge to the nearest vertical presentation, master, or layout guide at or left of that edge, so the shape never moves right; the slide edge is the fallback. Right-click for explicit selection, slide, first-anchor, last-anchor, and collision modes. Each action creates a PowerPoint Undo entry.

### Align Right

**Command ID:** `BtnAlignRight` · **Type:** Command · **Maturity:** Stable

Multi-select: aligns right edges to the configured first/last anchor. Single select: snaps to the nearest right-side vertical guide or slide edge.

**How to use**

With multiple selected shapes, keeps the anchor object configured in Settings (**First selected** or **Last selected**) fixed and aligns every other right edge to it. With one shape, snaps the right edge to the nearest vertical presentation, master, or layout guide at or right of that edge, so the shape never moves left; the slide edge is the fallback. Right-click for explicit selection, slide, first-anchor, last-anchor, and collision modes. Each action creates a PowerPoint Undo entry.

### Align Top

**Command ID:** `BtnAlignTop` · **Type:** Command · **Maturity:** Stable

Multi-select: aligns top edges to the configured first/last anchor. Single select: snaps to the nearest upper horizontal guide or slide edge.

**How to use**

With multiple selected shapes, keeps the anchor object configured in Settings (**First selected** or **Last selected**) fixed and aligns every other top edge to it. With one shape, snaps the top edge to the nearest horizontal presentation, master, or layout guide above it; the slide edge is the fallback. Right-click for explicit selection, slide, first-anchor, last-anchor, and collision modes. Each action creates a PowerPoint Undo entry.

### Analyze Slide Media Size

**Command ID:** `BtnSlideSize` · **Type:** Command · **Maturity:** Stable

Estimates the rendered PNG footprint of pictures, media, and embedded OLE objects on each slide.

**How to use**

Exports render previews of pictures, media, and embedded OLE objects to an isolated temporary folder, then reports an approximation of relative slide heaviness. Values are rendered PNG estimates, not the compressed bytes stored inside the PPTX. Right-click to run the analysis or directly remove all existing analysis labels.

### Copy Shape Adjustments

**Command ID:** `BtnAlignShapeAdjustments` · **Type:** Command · **Maturity:** Stable

Copies yellow-diamond adjustment values from the configured anchor to compatible selected shapes.

**How to use**

Uses the first or last selected anchor configured in Settings. Right-click to use the first or last selected shape explicitly without changing Settings. Every other selected shape with the same PowerPoint shape type and the same number of yellow-diamond adjustments receives the anchor's exact adjustment values. Size, position, rotation, text, and formatting remain unchanged; incompatible shapes are skipped and write failures roll back all copied values.

### Decrease Horizontal Spacing

**Command ID:** `BtnDecreaseHorizontalSpacing` · **Type:** Command · **Maturity:** Stable

Decreases the horizontal gap between selected shapes by the configured spacing increment.

**How to use**

Select two or more shapes, then reduce the gaps between them horizontally. Right-click any incremental spacing button to increase or decrease horizontal or vertical spacing directly. Configure the increment in Settings to match your layout system.

### Decrease Vertical Spacing

**Command ID:** `BtnDecreaseVerticalSpacing` · **Type:** Command · **Maturity:** Stable

Decreases the vertical gap between selected shapes by the configured spacing increment.

**How to use**

Select two or more shapes, then reduce the gaps between them vertically. Right-click any incremental spacing button to increase or decrease horizontal or vertical spacing directly. Configure the increment in Settings to match your layout system.

### Distribute Horizontally

**Command ID:** `BtnDistributeHorizontal` · **Type:** Command · **Maturity:** Stable

Distributes selected shapes horizontally within the selection or across the slide.

**How to use**

The primary action spaces three or more shapes evenly between the outer shapes; with exactly two selected shapes it spreads them across the slide instead, because there is no interior gap to distribute. Right-click to choose the across-slide route explicitly or to align their vertical middles before distribution. Each action creates one Undo entry.

### Distribute Vertically

**Command ID:** `BtnDistributeVertical` · **Type:** Command · **Maturity:** Stable

Distributes selected shapes vertically within the selection or across the slide.

**How to use**

The primary action spaces three or more shapes evenly between the outer shapes; with exactly two selected shapes it spreads them across the slide instead, because there is no interior gap to distribute. Right-click to choose the across-slide route explicitly or to align their horizontal centers before distribution. Each action creates one Undo entry.

### Increase Horizontal Spacing

**Command ID:** `BtnIncreaseHorizontalSpacing` · **Type:** Command · **Maturity:** Stable

Increases the horizontal gap between selected shapes by the configured spacing increment.

**How to use**

Select two or more shapes, then increase the gaps between them horizontally. Right-click any incremental spacing button to increase or decrease horizontal or vertical spacing directly. Configure the increment in Settings to match your layout system.

### Increase Vertical Spacing

**Command ID:** `BtnIncreaseVerticalSpacing` · **Type:** Command · **Maturity:** Stable

Increases the vertical gap between selected shapes by the configured spacing increment.

**How to use**

Select two or more shapes, then increase the gaps between them vertically. Right-click any incremental spacing button to increase or decrease horizontal or vertical spacing directly. Configure the increment in Settings to match your layout system.

### Lock/Unlock Aspect Ratio

**Command ID:** `BtnLockDimensions` · **Type:** Toggle · **Maturity:** Stable

Locks or unlocks the aspect ratio of selected shapes while still allowing proportional resizing.

**How to use**

Locks the width-to-height ratio of all selected shapes. Shapes can still be moved and resized proportionally. If the selection has mixed states, clicking locks all selected shapes; clicking again unlocks all. Right-click to lock or unlock every slide shape in the presentation.

### Match Height

**Command ID:** `BtnMatchHeight` · **Type:** Command · **Maturity:** Stable

Matches selected shapes to the height of the configured first or last selected anchor while preserving width and position.

**How to use**

Keeps the first or last selected anchor configured in Settings fixed and resizes the remaining shapes to its exact height. Right-click any Match button to choose width, height, or both dimensions directly. With exactly two selected shapes, **Match Height and Swap Positions** first matches the non-anchor height and then exchanges both center positions in one Undo step. Width and existing aspect-ratio lock state are preserved.

### Match Size

**Command ID:** `BtnMatchSize` · **Type:** Command · **Maturity:** Stable

Forces all selected objects to match the exact Width and Height of the anchor object.

**How to use**

Resizes every non-anchor shape to the exact width and height of the first or last selected anchor configured in Settings. Right-click any Match Size, Match Width, or Match Height button to choose width, height, or both dimensions directly. Existing aspect-ratio lock states are preserved, and the operation can be undone as one PowerPoint step.

### Match Width

**Command ID:** `BtnMatchWidth` · **Type:** Command · **Maturity:** Stable

Standardizes the width of all selected objects to match the anchor object.

**How to use**

Resizes every non-anchor shape to the exact width of the first or last selected anchor configured in Settings. Right-click any Match button to choose width, height, or both dimensions directly. Height, position, and existing aspect-ratio lock state are preserved.

### Remove Horizontal Gaps

**Command ID:** `BtnRemoveHorizontalGaps` · **Type:** Command · **Maturity:** Stable

Removes horizontal gaps between selected shapes.

**How to use**

Collapses all horizontal spacing between selected shapes to zero, sliding them together with no gap on the horizontal axis.

### Remove Vertical Gaps

**Command ID:** `BtnRemoveVerticalGaps` · **Type:** Command · **Maturity:** Stable

Removes vertical gaps between selected shapes.

**How to use**

Select at least two shapes. The command removes the vertical gaps while keeping the first selected shape in place and preserving the top-to-bottom order. Right-click the sidebar button and choose **Remove Vertical Gaps & Center** to center the packed group within its original vertical bounds instead.

### Scale Selected Shapes

**Command ID:** `BtnMagicResizer` · **Type:** Toggle · **Maturity:** Stable

Scales every selected shape by the same factor, with a shared anchor setting and optional text, line, margin, and bullet scaling.

**How to use**

Select one or more shapes — a group is a valid target and scales as one unit — then scale each one by the same percentage or enter target dimensions derived from the first selected shape. Choose which point on every shape remains fixed. Font size, line weight, text margins, and bullet size can be scaled with the shapes. The command keeps each shape's chosen anchor in place rather than scaling the complete selection bounding box.

### Shape Size Panel

**Command ID:** `BtnSizePanel` · **Type:** Widget · **Maturity:** Stable

Displays and edits exact width and height for the selected shapes using centimeter or point input.

**How to use**

The embedded sidebar fields show a shared width or height for the selected shapes and show **Mixed** when that dimension differs. Enter a positive value with `cm`, with `pt`, or without a suffix for centimeters to apply one exact dimension to every selected shape. The arrow buttons adjust each shape by 0.1 cm. Existing aspect-ratio locks remain active, Enter and focus loss do not apply the same edit twice, and each change creates one Undo entry with rollback on errors. Right-click to focus and select the width or height field directly.

### Stack Above

**Command ID:** `BtnDockTop` · **Type:** Command · **Maturity:** Stable

Moves selected shapes above the anchor shape, stacking them vertically.

**How to use**

Keeps the first or last selected anchor object configured in Settings fixed, then places every other selected shape directly above it in selection order with zero vertical gap. Right-click any Stack button to choose left, right, above, or below directly. Horizontal positions, sizes, and formatting remain unchanged.

### Stack Below

**Command ID:** `BtnDockBottom` · **Type:** Command · **Maturity:** Stable

Moves selected shapes below the anchor shape, stacking them vertically.

**How to use**

Keeps the first or last selected anchor object configured in Settings fixed, then places every other selected shape directly below it in selection order with zero vertical gap. Right-click any Stack button to choose left, right, above, or below directly. Horizontal positions, sizes, and formatting remain unchanged.

### Stack Left

**Command ID:** `BtnDockLeft` · **Type:** Command · **Maturity:** Stable

Moves selected shapes to the left of the anchor object, stacking them horizontally.

**How to use**

Keeps the first or last selected anchor object configured in Settings fixed, then places every other selected shape directly to its left in selection order with zero horizontal gap. Right-click any Stack button to choose left, right, above, or below directly. Vertical positions, sizes, and formatting remain unchanged.

### Stack Right

**Command ID:** `BtnDockRight` · **Type:** Command · **Maturity:** Stable

Moves selected shapes to the right side of the anchor shape, stacking them horizontally.

**How to use**

Keeps the first or last selected anchor object configured in Settings fixed, then places every other selected shape directly to its right in selection order with zero horizontal gap. Right-click any Stack button to choose left, right, above, or below directly. Vertical positions, sizes, and formatting remain unchanged.

### Stretch Bottom

**Command ID:** `BtnStretchBottom` · **Type:** Command · **Maturity:** Stable

Extends shapes downward across the gap until they touch the anchor's top edge.

**How to use**

Keeps the first or last selected anchor configured in Settings fixed. Shapes entirely above it extend their bottom edge to the anchor's top edge while their original top edge, width, horizontal position, and aspect-ratio lock state remain unchanged. Overlapping shapes and shapes on the other side are skipped. Right-click any Stretch button to choose a gap direction or move a specific edge to the anchor.

### Stretch Bottom Edge

**Command ID:** `BtnStretchBottomEdge` · **Type:** Command · **Maturity:** Stable

Resizes selected shapes so their bottom edges match the configured anchor's bottom edge.

**How to use**

Keeps the first or last selected anchor configured in Settings fixed. Each eligible shape keeps its top edge fixed and stretches its bottom edge to the anchor's bottom edge. Shapes whose top edge is at or beyond the target, and shapes already aligned to it, are skipped to avoid inversion and no-op changes. Right-click any Stretch button to choose a gap direction or move a specific edge to the anchor.

### Stretch Left

**Command ID:** `BtnStretchLeft` · **Type:** Command · **Maturity:** Stable

Extends shapes leftward across the gap until they touch the anchor's right edge.

**How to use**

Keeps the first or last selected anchor configured in Settings fixed. Shapes entirely to its right extend their left edge to the anchor's right edge while their original right edge, height, position on the other axis, and aspect-ratio lock state remain unchanged. Overlapping shapes and shapes on the other side are skipped. Right-click any Stretch button to choose a gap direction or move a specific edge to the anchor.

### Stretch Left Edge

**Command ID:** `BtnStretchLeftEdge` · **Type:** Command · **Maturity:** Stable

Resizes selected shapes so their left edges match the anchor's left edge.

**How to use**

Keeps the first or last selected anchor configured in Settings fixed. Each other shape whose right edge lies beyond the anchor's left edge is resized from the left until both left edges match; its original right edge, height, vertical position, and aspect-ratio lock state remain unchanged. Shapes entirely left of the target edge are skipped to prevent inversion. Right-click any Stretch button to choose a gap direction or move a specific edge to the anchor.

### Stretch Right

**Command ID:** `BtnStretchRight` · **Type:** Command · **Maturity:** Stable

Extends shapes rightward across the gap until they touch the anchor's left edge.

**How to use**

Keeps the first or last selected anchor configured in Settings fixed. Shapes entirely to its left extend their right edge to the anchor's left edge while their original left edge, height, position on the other axis, and aspect-ratio lock state remain unchanged. Overlapping shapes and shapes on the other side are skipped. Right-click any Stretch button to choose a gap direction or move a specific edge to the anchor.

### Stretch Right Edge

**Command ID:** `BtnStretchRightEdge` · **Type:** Command · **Maturity:** Stable

Resizes selected shapes so their right edges match the configured anchor's right edge.

**How to use**

Keeps the first or last selected anchor configured in Settings fixed. Each eligible shape keeps its left edge fixed and stretches its right edge to the anchor's right edge. Shapes whose left edge is at or beyond the target, and shapes already aligned to it, are skipped to avoid inversion and no-op changes. Right-click any Stretch button to choose a gap direction or move a specific edge to the anchor.

### Stretch Top

**Command ID:** `BtnStretchTop` · **Type:** Command · **Maturity:** Stable

Extends shapes upward across the gap until they touch the anchor's bottom edge.

**How to use**

Keeps the first or last selected anchor configured in Settings fixed. Shapes entirely below it extend their top edge to the anchor's bottom edge while their original bottom edge, width, horizontal position, and aspect-ratio lock state remain unchanged. Overlapping shapes and shapes on the other side are skipped. Right-click any Stretch button to choose a gap direction or move a specific edge to the anchor.

### Stretch Top Edge

**Command ID:** `BtnStretchTopEdge` · **Type:** Command · **Maturity:** Stable

Resizes selected shapes so their top edges match the anchor's top edge.

**How to use**

Keeps the first or last selected anchor configured in Settings fixed. Each other shape whose bottom edge lies below the anchor's top edge is resized from the top until both top edges match; its original bottom edge, width, horizontal position, and aspect-ratio lock state remain unchanged. Shapes entirely above the target edge are skipped to prevent inversion. Right-click any Stretch button to choose a gap direction or move a specific edge to the anchor.

## Arrange and organize

### Arrange Grid

**Command ID:** `BtnArrangeGrid` · **Type:** Command · **Maturity:** Stable

Opens a grid layout dialog for arranging selected shapes with your chosen rows, columns, and spacing.

**How to use**

Select two or more shapes, then choose positive row and column counts and a non-negative spacing. Press Enter to apply valid values or Escape to close without arranging. Right-click for direct two-, three-, four-, or five-column grids with automatic row calculation and 10-point spacing. Shapes follow selection order from the selection's top-left position; each column and row accommodates its largest shape. The operation can be undone as one step.

### Arrange Shapes Geometrically

**Command ID:** `BtnArrangePro` · **Type:** Command · **Maturity:** Stable

Opens a live, collision-aware geometric arrangement palette for selected shapes.

**How to use**

Arrange selected shapes around a circle, configurable arc, triangle, square, pentagon, octagon, or star. Adjust radius, start angle, center, object order, clockwise direction, and optional inward rotation. **First Shape** or **Last Shape** keeps that shape unchanged as the center anchor and arranges the remaining shapes around it. **Avoid Overlap** increases the effective radius until the arranged shapes and optional anchor no longer intersect. With one selected shape, Instances creates tagged duplicates; in anchor mode the source remains in the center and the requested number of duplicates is arranged around it. Fill Shape is available for circle, triangle, and square. The palette is limited to one PowerPoint-owned window.

### Arrange Shapes in Circle

**Command ID:** `BtnArrangeCircle` · **Type:** Command · **Maturity:** Stable

Arranges selected shapes around a circular path. Right-click to choose another enclosing shape.

**How to use**

Select at least two shapes and click the feature to distribute them around a circle. Right-click to arrange the selection in a triangle, square, star, octagon, or pentagon instead.

### Arrange Shapes in Pattern

**Command ID:** `BtnArrangeInShape` · **Type:** Command · **Maturity:** Stable

Arranges selected shapes in a circular, triangular, or other pattern.

**How to use**

Distributes selected shapes into a geometric pattern within their total bounding area.

**Right-Click Options:**
* **Patterns**: Choose between Circle, Triangle, Square, Star, Octagon, or Pentagon.
* **Pro Mode**: Opens a dialog for fine-tuned control over radius, start angle, and auto-rotation towards center.

### Bring Forward

**Command ID:** `BtnBringForward` · **Type:** Command · **Maturity:** Stable

Brings the selected shapes forward by one layer.

**How to use**

Moves the selection forward by one position in the stacking order. Right-click variants:

- **Rank Selected Shapes (First Selected on Top)** performs the same one-step move and puts the first selected shape on top.
- **Place Directly Above Last Selected** uses the last selected shape as the reference.
- **Place Directly Below Last Selected** uses the last selected shape as the reference.

Relative placement preserves the movers' current internal layer order and the relative order of unselected shapes.

### Bring to Front

**Command ID:** `BtnBringToFront` · **Type:** Command · **Maturity:** Stable

Brings the selected shapes to the absolute front layer.

**How to use**

Moves the selected shapes above every other object on the slide. Right-click variants:

- **Rank Selected Shapes (First Selected on Top)** makes the first selected shape topmost.
- **Place Directly Above Last Selected** uses the last selected shape as the reference.
- **Place Directly Below Last Selected** uses the last selected shape as the reference.

Relative placement preserves the movers' current internal layer order. Unselected shapes keep their relative order.

### Connect Shapes

**Command ID:** `BtnObjectConnector` · **Type:** Command · **Maturity:** Stable

Creates a filled connector shape between two selected objects. Auto-detects closest sides.

**How to use**

Creates a filled connector polygon between the first two selected shapes. The primary action detects the nearest sides and follows supported shape silhouettes. Right-click to use automatic smart routing or force right-to-left, left-to-right, bottom-to-top, or top-to-bottom attachment sides.

### Copy XY-Coordinates

**Command ID:** `BtnCopyXY` · **Type:** Command · **Maturity:** Stable

Copies the exact top-left X/Y position from one selected shape for Paste XY-Coordinates.

**How to use**

Select exactly one shape, then copy its Left and Top position as a versioned openEfficiency X/Y payload for Paste XY-Coordinates. This command never copies Width or Height. Paste XY-Coordinates also accepts legacy two- or four-value X/Y payloads, but right-click Copy XY-Coordinates clears only a valid current openEfficiency payload and never unrelated numeric clipboard text. Use Copy/Paste Position & Size when you need to transfer dimensions.

### Copy/Paste Position & Size

**Command ID:** `BtnPositionPainter` · **Type:** Command · **Maturity:** Stable

Copies the exact position and dimensions of one shape and pastes chosen properties onto other shapes in the same presentation.

**How to use**

Select exactly one source shape to copy its position and size and open the tool. Choose Left, Top, Width, and/or Height, select one or more target shapes in the same presentation, and click Paste. Right-click presets the tool for position only, size only, or both. Copy Again replaces the source values with the current single-shape selection. Target aspect-ratio locks are preserved, partial failures are rolled back, and each successful Paste is one PowerPoint Undo step.

### Create Motion Path

**Command ID:** `BtnCreateMotionPath` · **Type:** Command · **Maturity:** Stable

Adds an on-click motion path from the first selected shape's center to the second selected shape's center.

**How to use**

Select exactly two shapes. The primary action creates an on-click motion path from the first selected shape's center to the second. Right-click can reverse the direction or create both directions in one Undo step. Existing animations remain unchanged, the shapes do not move while editing, and incomplete effects are removed if creation fails.

### Flip Horizontal

**Command ID:** `BtnFlipHorizontal` · **Type:** Command · **Maturity:** Stable

Flips every selected shape horizontally around its own center.

**How to use**

Mirrors every selected shape from left to right around its own center. Right-click either Flip button to choose horizontal or vertical mirroring directly. Position and size stay unchanged, and the complete selection can be restored with one PowerPoint Undo step.

### Flip Vertical

**Command ID:** `BtnFlipVertical` · **Type:** Command · **Maturity:** Stable

Flips every selected shape vertically around its own center.

**How to use**

Mirrors every selected shape from top to bottom around its own center. Right-click either Flip button to choose horizontal or vertical mirroring directly. Position and size stay unchanged, and the complete selection can be restored with one PowerPoint Undo step.

### Hide Selected

**Command ID:** `BtnHideSelected` · **Type:** Command · **Maturity:** Stable

Hides all currently selected shapes on the slide. Use 'Show Hidden' to make them visible again.

**How to use**

Temporarily hides visible selected shapes without deleting them and marks them for safe restoration. Right-click to hide the selection, restore only shapes hidden by openEfficiency, or deliberately reveal every hidden shape on the slide after confirmation. Unrelated hidden shapes stay hidden during the safe restoration action.

### Hide/Show Master Objects

**Command ID:** `BtnHideMasterObjects` · **Type:** Command · **Maturity:** Stable

Hides or restores master shapes and the master background on the selected slides or current slide.

**How to use**

Hides master shapes and the master background on the selected slides, or on the current slide when no slide range is selected. A mixed selection is normalized by hiding all selected slides. The next click restores each slide's original master-shape and master-background settings. Right-click to hide, show, or toggle explicitly. Slide-level content is not changed.

### Lock/Unlock Shapes

**Command ID:** `BtnShapeLock` · **Type:** Toggle · **Maturity:** Stable

Locks or unlocks the selected shapes to prevent accidental edits.

**How to use**

Locks or unlocks the selected shapes using PowerPoint's native shape lock. If the selection contains both states, left-click locks every selected shape; left-click again unlocks them.

**Right-click options:**
* **Toggle Lock All Shapes**: Lock all shapes when any shape is unlocked; unlock all when every shape is locked.
* **Lock All Shapes**: Lock all shapes on every slide in the presentation.
* **Unlock All Shapes**: Unlock all shapes on every slide in the presentation.

Native shape locking requires a PowerPoint version that supports it.

### Paste XY-Coordinates

**Command ID:** `BtnPasteXY` · **Type:** Command · **Maturity:** Stable

Applies a copied top-left X/Y position to one selected shape without changing its size.

**How to use**

Select exactly one target shape, then paste the Left and Top position copied by Copy XY-Coordinates. Width, Height, rotation, text, and formatting stay unchanged. Current payloads are versioned and openEfficiency-owned; legacy two- and four-value X/Y payloads remain paste-compatible, with four-value Width and Height ignored. Right-click Paste XY-Coordinates clears only a valid current openEfficiency payload, never unrelated numeric clipboard text. Use Copy/Paste Position & Size when you need to transfer dimensions.

### Rotate Shapes

**Command ID:** `BtnRotateShape` · **Type:** Command · **Maturity:** Stable

Rotates every selected shape 90° clockwise; the context menu provides 90° left and 45° steps.

**How to use**

Click to rotate every selected shape 90° clockwise. Right-click for 90° left, 45° right, or 45° left. Each shape rotates around its own center, and the complete selection can be restored with one PowerPoint Undo step.

### Select Same Type

**Command ID:** `BtnSelectSameType` · **Type:** Command · **Maturity:** Stable

Selects shapes on the current slide that match the first selected shape.

**How to use**

Left-click to select shapes with the same PowerPoint shape type as the first selected shape. Right-click to match by fill color, line color, fill and line color, size, width, height, or type and formatting. The reference shape remains part of the resulting selection.

### Send Backward

**Command ID:** `BtnSendBackward` · **Type:** Command · **Maturity:** Stable

Sends the selected shapes backward by one layer.

**How to use**

Moves the selection backward by one position in the stacking order. Right-click variants:

- **Rank Selected Shapes (First Selected at Bottom)** performs the same one-step move and puts the first selected shape at the bottom.
- **Place Directly Above Last Selected** uses the last selected shape as the reference.
- **Place Directly Below Last Selected** uses the last selected shape as the reference.

Relative placement preserves the movers' current internal layer order and the relative order of unselected shapes.

### Send to Back

**Command ID:** `BtnSendToBack` · **Type:** Command · **Maturity:** Stable

Sends the selected shapes to the absolute back layer.

**How to use**

Moves the selected shapes behind every other object on the slide. Right-click variants:

- **Rank Selected Shapes (First Selected at Bottom)** makes the first selected shape bottommost.
- **Place Directly Above Last Selected** uses the last selected shape as the reference.
- **Place Directly Below Last Selected** uses the last selected shape as the reference.

Relative placement preserves the movers' current internal layer order. Unselected shapes keep their relative order.

### Show Hidden

**Command ID:** `BtnShowHidden` · **Type:** Command · **Maturity:** Stable

Restores shapes hidden by openEfficiency on the current slide without revealing unrelated hidden shapes.

**How to use**

Restores shapes previously hidden with Hide Selected at their original position and Z-order. Shapes hidden through PowerPoint or another workflow remain hidden. Right-click to choose the safe openEfficiency-only restoration or the confirmed recovery action that reveals every hidden shape on the current slide.

### Simplify Freeform Nodes

**Command ID:** `BtnOptimizeFreeForm` · **Type:** Command · **Maturity:** Stable

Removes adjacent freeform nodes that are less than five points apart.

**How to use**

Cleans selected freeform shapes, including freeforms inside selected groups, by removing adjacent nodes whose distance is less than 5 pt. Endpoints are retained and the shape object is edited in place. Curve or subpath nodes that PowerPoint does not allow to be deleted are skipped and logged. The command does not perform global smoothing or mathematical outlier detection.

### Snap Rotation to 90°

**Command ID:** `BtnRectifyRotation` · **Type:** Command · **Maturity:** Stable

Snaps every selected shape to the nearest 90° orientation and straightens near-horizontal or near-vertical lines.

**How to use**

Snaps each selected shape to its nearest 0°, 90°, 180°, or 270° orientation. Lines and connectors are straightened to the nearer horizontal or vertical axis. Right-click to set the complete selection explicitly to 0°, 90°, 180°, or 270°. Each operation can be restored with one PowerPoint Undo step.

### Snap Shape to Object

**Command ID:** `BtnSnapShape` · **Type:** Command · **Maturity:** Stable

Moves the first of exactly two selected shapes flush to the nearest side of the second shape.

**How to use**

Select exactly two shapes. The first selected shape moves flush to the nearest side of the second and is centered on the perpendicular axis; size and formatting are preserved. Right-click for an explicit left, right, above, or below direction.

### Snap to Grid

**Command ID:** `BtnSnapToGrid` · **Type:** Toggle · **Maturity:** Stable

Toggles snap-to-grid functionality for precise shape positioning.

**How to use**

Enables session-level grid snapping for one selected shape while it is dragged in PowerPoint. The shape's top-left corner is placed on the nearest 50-point grid intersection and temporary guides are removed after the drag. Enabling this mode disables Snap to Objects to prevent conflicting mouse hooks. Right-click to enable Grid or Object snapping explicitly, or disable both snapping modes.

### Snap to Objects

**Command ID:** `BtnSnapToObjects` · **Type:** Toggle · **Maturity:** Stable

Toggles snap-to-objects functionality for precise shape alignment.

**How to use**

Enables session-level smart snapping for one selected shape while it is dragged in PowerPoint. Visible shapes on the current slide provide edge and center targets; temporary guides disappear after the drag. Enabling this mode disables Snap to Grid to prevent conflicting mouse hooks. Right-click to enable Object or Grid snapping explicitly, or disable both snapping modes.

### Swap Positions

**Command ID:** `BtnSwapPositions` · **Type:** Command · **Maturity:** Stable

Exchanges the X/Y coordinates of two selected shapes. Ideal for reorganization.

**How to use**

Select exactly two shapes to exchange their positions. The standard action swaps their centers. Right-click to instead exchange matching top, side, corner, or bottom anchor points. Shape size, rotation, and formatting remain unchanged.

### Sync Shapes Across Slides

**Command ID:** `BtnSyncObjects` · **Type:** Command · **Maturity:** Stable

Opens dialog to distribute and sync shapes across multiple slides.

**How to use**

Distributes selected shapes to chosen slides, then synchronizes tagged copies. Right-click can synchronize format, position, content, or all properties directly from the one selected tagged source shape; target-slide distribution remains in the full dialog.

### Toggle Rounded Corners

**Command ID:** `BtnSmartCorners` · **Type:** Command · **Maturity:** Stable

Synchronizes selected native rectangles between square and uniformly rounded corners.

**How to use**

Select one or more native Rectangle or Rounded Rectangle shapes. The primary action toggles all compatible shapes between square and uniformly rounded. Right-click to force Rounded, force Square, or Toggle. Incompatible shapes are skipped. Shape objects, text, formatting, tags, links, animations, and Z-order are preserved because the native AutoShape type changes in place.

## Text and formatting

### Adjust Transparency

**Command ID:** `BtnTransparency` · **Type:** Toggle · **Maturity:** Stable

Open a floating toolbar to adjust the transparency of selected shapes.

**How to use**

Opens a floating toolbar for changing the transparency of the selected shapes. Use it to make fills, lines, or entire objects more subtle while keeping the underlying layout intact. Select one or more shapes first, then choose the desired transparency in the toolbar.

### Align Text

**Command ID:** `BtnAlignText` · **Type:** Command · **Maturity:** Stable

Opens PowerPoint's text-alignment gallery or directly aligns text inside selected shapes.

**How to use**

The primary action opens PowerPoint's native text-alignment gallery. Right-click to set horizontal alignment, vertical anchoring, or one of the centered combinations directly. Direct actions affect the active text box or every compatible selected shape, skip shapes without a text frame, and create one PowerPoint Undo entry.

### Bold

**Command ID:** `BtnFormatBold` · **Type:** Command · **Maturity:** Stable

Toggles PowerPoint bold formatting for the current text selection.

**How to use**

The primary action runs PowerPoint's native Bold toggle. Right-click to explicitly apply, remove, or toggle Bold for selected characters or all text in selected text-bearing shapes. Shapes without text are skipped and each explicit action supports Undo.

### Change Case

**Command ID:** `BtnChangeCase` · **Type:** Command · **Maturity:** Stable

Opens PowerPoint's Change Case menu or applies a selected case directly.

**How to use**

The primary action opens PowerPoint's native Change Case menu. Right-click to directly apply Sentence case, lowercase, UPPERCASE, Capitalize Each Word, or tOGGLE cASE to selected characters or all text in selected text-bearing shapes. PowerPoint controls language-specific casing; non-letter characters and formatting remain unchanged, and direct choices create one Undo entry.

### Character Spacing

**Command ID:** `BtnCharacterSpacing` · **Type:** Command · **Maturity:** Stable

Opens PowerPoint's spacing gallery or the custom spacing dialog for selected text.

**How to use**

The primary action opens PowerPoint's native Character Spacing gallery, whose presets (Very Tight, Tight, Normal, Loose, Very Loose) apply to the selection. Right-click **More Spacing...** to open PowerPoint's native Font dialog for exact values. PowerPoint's Font object model exposes no spacing property, so oPenEfficiency cannot apply exact presets directly: the native gallery and dialog are the supported contract.

### Color Eraser

**Command ID:** `BtnTransparentColor` · **Type:** Command · **Maturity:** Stable

Rasterizes one selected shape and makes pixels matching a picked screen color transparent.

**How to use**

Select exactly one shape, pick a color from the screen, adjust the tolerance, and erase matching pixels. The selected shape is exported and replaced with a PNG, so editable text, vector geometry, animations, links, and other shape-specific properties are flattened.

### Color Picker

**Command ID:** `BtnPickColor` · **Type:** Command · **Maturity:** Stable

Samples an exact screen pixel and applies its absolute RGB value to compatible properties of the selected shapes.

**How to use**

Select one or more shapes, then click the picker and choose any visible screen pixel. The primary action changes fill foreground colors; right-click provides line and text targets. Fill styles such as gradients are preserved, text applies to the complete text of each compatible selected shape, Escape cancels without changes, and each successful multi-shape application is one PowerPoint Undo step.

### Decrease Font Size

**Command ID:** `BtnDecreaseFontSize` · **Type:** Command · **Maturity:** Stable

Moves selected text to PowerPoint's next smaller font-size step.

**How to use**

Uses PowerPoint's native Decrease Font Size command. Right-click to move one, two, or five steps on PowerPoint's built-in size ladder. It affects selected characters or text in selected text-bearing shapes; mixed font sizes remain relative.

### Delete Text

**Command ID:** `BtnDeleteText` · **Type:** Command · **Maturity:** Stable

Clears text from selected shapes or selected table cells while preserving the shapes themselves.

**How to use**

Clears text from selected text-bearing shapes, selected table cells, and supported children of selected groups while keeping the shapes. Right-click to delete all text, only highlighted characters, or the paragraphs intersecting the active text selection. **Delete Highlighted Text** and **Delete Selected Paragraphs** are available only for an actual text selection. You can also open **Replace Selected Text...**. Each deletion creates one PowerPoint Undo entry.

### Fit Shape to Text

**Command ID:** `BtnFitFormToText` · **Type:** Command · **Maturity:** Stable

Fits selected text-bearing shapes tightly around their text, including supported shapes inside groups.

**How to use**

For each selected text-bearing shape, including supported children of selected groups, disables word wrapping and lets PowerPoint resize the shape to fit its text. The primary action also removes all four text margins. Right-click to choose fitting with zero margins or while preserving the existing margins. Shapes without text are skipped and the complete action creates one PowerPoint Undo entry.

### Font Panel

**Command ID:** `BtnFontPanel` · **Type:** Widget · **Maturity:** Stable

Shows embedded controls for changing the font family and size of the active text selection or selected text-bearing shapes.

**How to use**

Choose or type a font family and font size for the active text selection or all selected text-bearing shapes. Mixed selections show an empty field. Enter and focus loss apply a typed value only once; successful changes create one PowerPoint Undo entry. Right-click to focus and select either field directly.

### Format Numbers

**Command ID:** `BtnFormatNumbers` · **Type:** Command · **Maturity:** Stable

Reformats detected numbers with configurable grouping, decimal separators, and currency placement.

**How to use**

Choose standard comma, period, space, apostrophe, or Indian grouping and optionally reposition, replace, or remove a detected currency symbol. Right-click for direct US, German, French-space, or Swiss-apostrophe separator presets. The command affects the active text selection, selected shapes, selected slides, or the current slide. Other text and formatting remain in place, and successful changes create one Undo entry.

### Increase Font Size

**Command ID:** `BtnIncreaseFontSize` · **Type:** Command · **Maturity:** Stable

Moves selected text to PowerPoint's next larger font-size step.

**How to use**

Uses PowerPoint's native Increase Font Size command. Right-click to move one, two, or five steps on PowerPoint's built-in size ladder. It affects selected characters or text in selected text-bearing shapes; mixed font sizes remain relative.

### Insert Special Characters

**Command ID:** `BtnSpecialChars` · **Type:** Command · **Maturity:** Stable

Opens a persistent picker for inserting symbols, currencies, operators, punctuation, and language-specific characters.

**How to use**

Place the cursor in text or select exactly one text-bearing shape, then choose a category and click a character. It is inserted after the active cursor or selection, or appended to the selected shape. The context menu opens the picker directly at a symbol or language category; each insertion creates one PowerPoint Undo entry.

### Insert Text

**Command ID:** `BtnInsertText` · **Type:** Command · **Maturity:** Stable

Prompts for text and inserts it at the current text cursor or at the end of one selected text-bearing shape.

**How to use**

Enter a text snippet, then insert it after the active text cursor or selection. Right-click to insert a line break, bullet, en dash, or em dash directly. With exactly one text-bearing shape selected, the text is appended to that shape. Cancel and empty input leave the presentation unchanged; successful insertion creates one PowerPoint Undo entry.

### Italic

**Command ID:** `BtnFormatItalic` · **Type:** Command · **Maturity:** Stable

Toggles PowerPoint italic formatting for the current text selection.

**How to use**

The primary action runs PowerPoint's native Italic toggle. Right-click to explicitly apply, remove, or toggle Italic for selected characters or all text in selected text-bearing shapes. Shapes without text are skipped and each explicit action supports Undo.

### Open Format Shape Pane

**Command ID:** `BtnFormatShapeDialog` · **Type:** Command · **Maturity:** Stable

Opens the native PowerPoint 'Format Shape' pane to adjust fill, line, and effects.

**How to use**

Opens PowerPoint's native Format Shape interface for the current shape selection. Right-click can remove individual effects or all effects. Reset to Default Shape Formatting uses a newly inserted AutoShape of the same type as the theme-derived reference while preserving the selected shape's geometry, position, text, and text formatting. Reset is unavailable for objects without a meaningful same-type AutoShape default.

### Paragraph & Line Spacing

**Command ID:** `BtnParagraphDialog` · **Type:** Command · **Maturity:** Stable

Opens PowerPoint's native paragraph dialog or applies common spacing presets.

**How to use**

Select text or one or more text-bearing shapes, then open PowerPoint's native paragraph dialog to adjust alignment, indentation, line spacing, and paragraph spacing. Right-click applies single, 1.15, or 1.5 line spacing immediately, or removes paragraph spacing. Each preset preserves unrelated paragraph properties and creates one Undo entry.

### Pick Fill Color

**Command ID:** `BtnPickFillColor` · **Type:** Command · **Maturity:** Stable

Samples an exact screen pixel and applies its absolute RGB value to the fill foreground of selected shapes.

**How to use**

Select one or more shapes, then choose any visible screen pixel. The sampled absolute RGB value is applied to each compatible fill foreground while the existing fill style is preserved. Right-click can redirect the picker to fill, line, or text color. Escape cancels without changes; multi-shape changes use one Undo step and roll back together after an error.

### Pick Line Color

**Command ID:** `BtnPickLineColor` · **Type:** Command · **Maturity:** Stable

Samples an exact screen pixel and applies its absolute RGB value to outlines and line shapes.

**How to use**

Select one or more shapes or lines, then choose any visible screen pixel. The sampled absolute RGB value is applied to each compatible outline and makes hidden lines visible. Right-click can redirect the picker to fill, line, or text color. Escape cancels without changes; multi-shape changes use one Undo step and roll back together after an error.

### Pick Text Color

**Command ID:** `BtnPickTextColor` · **Type:** Command · **Maturity:** Stable

Samples an exact screen pixel and applies its absolute RGB value to all text in compatible selected shapes.

**How to use**

Select one or more text-capable shapes, then choose any visible screen pixel. The sampled absolute RGB value replaces the text color across each compatible selected shape; non-text shapes are skipped. Right-click can redirect the picker to fill, line, or text color. Escape cancels without changes, and partial failures restore captured rich-text formatting where possible.

### Replace Formatting

**Command ID:** `BtnReplaceFont` · **Type:** Command · **Maturity:** Stable

Replaces matching fonts, sizes, text colors, fills, and lines within a chosen slide scope.

**How to use**

Choose one or more format properties, define each find and replacement value, and apply them to the current slide, selected slides, or the entire presentation. Groups are traversed recursively. Matching is exact except for a 0.1 pt font-size tolerance. Invalid criteria are rejected before mutation; successful changes form one Undo entry, and an apply error triggers rollback of changes made by that run.

### Replace Text

**Command ID:** `BtnApplyTextTool` · **Type:** Command · **Maturity:** Stable

Opens a compact editor that replaces the active text selection, selected text shapes, or selected table cells.

**How to use**

Enter replacement text and apply it to the active text selection, every selected text-bearing shape, or the selected cells of a table. Selecting the table itself applies the text to every cell. **Take text from Master** uses the configured first- or last-selected master shape and is available only when that master can supply text. Right-click can use the first or last selected shape explicitly. Text replacement creates one PowerPoint Undo entry; it does not copy font or paragraph formatting.

### Set Spell Check Language

**Command ID:** `BtnSpellCheck` · **Type:** Command · **Maturity:** Stable

Sets the proofing language for selected shapes, selected slides, or the entire presentation.

**How to use**

Choose a proofing language and apply it to selected shapes, selected slides, or the entire presentation. Presentation scope also updates slide masters, custom layouts, and the presentation's default language; narrower scopes do not change that global default. Groups and table cells are included, and successful changes create one PowerPoint Undo entry.

### Shape Fill

**Command ID:** `BtnShapeFill` · **Type:** Command · **Maturity:** Stable

Opens a compact PowerPoint-style palette for theme-aware, standard, recent and custom shape fills.

**How to use**

Choose a theme color and tint, a standard or recent color, No Fill, More Fill Colors, or Eyedropper. The palette applies one fill operation to all eligible selected shapes and opening it never changes the document.

### Split Text Box by Paragraphs

**Command ID:** `BtnSplitByParagraphs` · **Type:** Command · **Maturity:** Stable

Splits one multi-paragraph text box into individually editable text boxes. Use the context menu to merge selected text boxes.

**How to use**

Select one text box with multiple paragraphs, then split it into individually editable text boxes. Right-click to omit empty paragraphs or retain them as empty shapes. The command preserves basic font, alignment, margin, and paragraph settings where PowerPoint permits and confirms before replacing the original. **Merge Selected Text Boxes** combines multiple text boxes in top-to-bottom, left-to-right order using the basic formatting of the first shape.

### Strikethrough

**Command ID:** `BtnFormatStrikethrough` · **Type:** Command · **Maturity:** Stable

Toggles PowerPoint strikethrough without changing bold formatting.

**How to use**

Runs PowerPoint's native Strikethrough toggle on the current text selection. PowerPoint's Font object model exposes no strikethrough property, so oPenEfficiency cannot explicitly apply or remove strikethrough: each click flips the selection's state and PowerPoint Undo reverts it. Right-click repeats the toggle.

### Subscript

**Command ID:** `BtnFormatSubscript` · **Type:** Command · **Maturity:** Stable

Toggles PowerPoint subscript for the current text selection.

**How to use**

The primary action runs PowerPoint's native Subscript toggle. Right-click to explicitly apply, remove, or toggle Subscript for selected characters or all text in selected text-bearing shapes. PowerPoint keeps it mutually exclusive with Superscript; each explicit action supports Undo.

### Superscript

**Command ID:** `BtnFormatSuperscript` · **Type:** Command · **Maturity:** Stable

Toggles PowerPoint superscript for the current text selection.

**How to use**

The primary action runs PowerPoint's native Superscript toggle. Right-click to explicitly apply, remove, or toggle Superscript for selected characters or all text in selected text-bearing shapes. PowerPoint keeps it mutually exclusive with Subscript; each explicit action supports Undo.

### Swap Text (Formatted)

**Command ID:** `BtnSwapTextFormatted` · **Type:** Command · **Maturity:** Stable

Swaps text and its character formatting between exactly two selected shapes.

**How to use**

Exchanges the complete text ranges of exactly two selected text-bearing shapes. Font runs, paragraph formatting, and other text formatting travel with the source text; shape geometry and shape formatting remain in place. The operation uses PowerPoint's clipboard, cleans up its temporary slide objects, attempts to restore both text ranges after an error, and creates one Undo entry.

### Swap Text (Plain)

**Command ID:** `BtnSwapTextPlain` · **Type:** Command · **Maturity:** Stable

Swaps plain text between two text shapes or selected table cells while retaining destination formatting.

**How to use**

Click to exchange the raw text strings of exactly two selected text-bearing shapes or exactly two selected table cells. The incoming text adopts the destination formatting; source character runs are not transferred. Right-click to shrink overflowing text after the swap, cycle plain text through three or more selected targets in selection order, or choose one paragraph from each of two targets to exchange. Invalid selections show a specific explanation. Every successful action creates one PowerPoint Undo entry and attempts rollback after an error.

### Text Direction

**Command ID:** `BtnTextDirection` · **Type:** Command · **Maturity:** Stable

Opens PowerPoint's text-direction gallery or directly changes text orientation.

**How to use**

The primary action opens PowerPoint's native Text Direction gallery. Right-click to apply Horizontal, Rotate all text 90°, Rotate all text 270°, or Stacked to the active text box or every compatible selected shape. Direct actions create one PowerPoint Undo entry.

### Text Format Painter

**Command ID:** `BtnTextFormatPainter` · **Type:** Command · **Maturity:** Stable

Copies basic text and text-frame formatting from the first selected shape to the remaining selected shapes without changing their text.

**How to use**

Select two or more text-bearing shapes. The first selected shape is the source; every other selected shape is a target. The primary action copies the source's first-character typography, paragraph alignment, vertical alignment, orientation, and text margins while retaining target text. Theme-linked text colors remain linked when PowerPoint exposes their theme value. Right-click **Apply Typography Only** to copy font family, size, color, and emphasis; **Apply Paragraphs & Bullets Only** to copy paragraph spacing, alignment, standard character or numbered bullets, and ruler indents by matching list level; or **Apply While Preserving Emphasis** to copy the basic font family, size, paragraph alignment, and text-frame settings without replacing each target's color, bold, italic, underline, subscript, or superscript emphasis. Picture bullets are left unchanged. Shapes without a text frame are skipped, the source must contain text, and each operation creates one PowerPoint Undo entry.

### Theme Color Palette (RGB)

**Command ID:** `BtnThemeColor` · **Type:** Command · **Maturity:** Stable

Opens a palette that resolves the active presentation theme to fixed RGB fill, line, and text colors.

**How to use**

Open the palette, then use the Fill, Border, or Text row to apply the active presentation's theme colors and tint variants as absolute RGB values. Right-click to open the floating palette directly at the Fill, Line, or Text colors, or show and hide the docked palette. The palette refreshes when the active presentation changes, so applied colors do not drift after a later theme change. Custom swatches are session-only fill colors.

### Translate Text

**Command ID:** `BtnTranslate` · **Type:** Command · **Maturity:** Stable

Opens PowerPoint's native translator or translates selected text through a configured external provider.

**How to use**

Click to open PowerPoint's native translation interface. Use the context menu to send the active text selection or selected text-bearing shapes to the configured DeepL, OpenAI, Claude, LibreTranslate, or free MyMemory service. External providers use HTTPS; a custom LibreTranslate endpoint may use HTTP only for a loopback server such as localhost. API translation replaces plain text and reapplies the common base font properties; mixed rich-text runs cannot be preserved exactly. The result is applied only if the source presentation is still active and creates one PowerPoint Undo entry.

### Underline

**Command ID:** `BtnFormatUnderline` · **Type:** Command · **Maturity:** Stable

Toggles PowerPoint underline formatting for the current text selection.

**How to use**

The primary action runs PowerPoint's native Underline toggle. Right-click to explicitly apply, remove, or toggle Underline for selected characters or all text in selected text-bearing shapes. Shapes without text are skipped and each explicit action supports Undo.

## Tables

### Align to Cell

**Command ID:** `BtnAlignToTableCell` · **Type:** Command · **Maturity:** Stable

Aligns selected non-table shapes to one of nine positions inside the table cell containing each shape's center.

**How to use**

Place one or more shapes over table cells, then choose Top/Center/Bottom combined with Left/Center/Right. The right-click menu exposes all nine positions directly in Top, Middle, and Bottom groups. Each shape is matched to the visible table cell containing its center and moved independently; table shapes themselves are ignored. The center action aligns both axes. Shapes whose centers are outside every table cell remain unchanged, and a successful action creates a PowerPoint Undo entry.

### Apply Table Heatmap

**Command ID:** `BtnTableHeatmap` · **Type:** Command · **Maturity:** Stable

Applies a row-, column-, or table-scoped color scale or icon prefix to numeric table cells.

**How to use**

Choose row, column, or complete-table normalization; one of 12 color presets or custom colors; fill, font color, or an icon prefix; and whether to exclude the first row or ignore blank cells. A live sample preview shows the active scheme, scope and data options before anything is applied. Right-click opens the dialog with row, column, or entire-table scope preselected. Non-empty non-numeric labels are always left unchanged. Icon mode replaces only a known leading heatmap icon and preserves the remaining cell text. If no numeric values are available nothing changes; successful application creates a PowerPoint Undo entry.

### Calculate Table Totals

**Command ID:** `BtnTableSum` · **Type:** Command · **Maturity:** Stable

Writes row and column totals into the last selected column and row while validating all input values first.

**How to use**

**How to use:** Click into a table and select the cell range that contains your numbers plus the row and/or column where the totals belong — for example the last column, the last row, or both. Then run Calculate Table Totals. Row totals are written into the last selected column, column totals into the last selected row, and the bottom-right cell receives the grand total. To calculate column totals only, use the context action instead of the primary command.

The selection needs at least two rows for column totals and at least two columns for row totals. Empty input cells are ignored; a non-empty value that cannot be parsed in the current, invariant, English, or German number format aborts the complete operation instead of being treated as zero. Outputs are prepared before any cell changes, destination formatting stays in place, and successful writes create a PowerPoint Undo entry.

### Convert Shapes to Table

**Command ID:** `BtnConvertToTable` · **Type:** Command · **Maturity:** Stable

Converts selected, uniformly aligned shapes into a PowerPoint table while retaining text and available basic cell formatting.

**How to use**

Select two or more non-overlapping shapes arranged in a uniform grid. Rectangles and rounded rectangles work especially well, but other shapes are accepted when you only need their text and available basic formatting in table cells. The command validates row and column alignment before it creates a table. The original shapes remain unchanged until conversion succeeds, then you can delete them, keep them visible, or discard the new table.

The conversion retains plain text plus available basic fill, line, font, alignment, and margin settings. Shapes without text create empty cells. Rich text, hyperlinks, gradients, patterns, effects, images, custom geometry, rotations, and gaps between shapes are not retained. Review the table before deleting the original shapes.

### Convert Table to Shapes

**Command ID:** `BtnConvertTableToShapes` · **Type:** Command · **Maturity:** Stable

Converts one PowerPoint table into a group of editable cell shapes while retaining basic text, fill, line, and alignment formatting.

**How to use**

Converts one selected native PowerPoint table into a group of independently editable cell shapes at the same slide coordinates. The primary action asks whether to delete, keep, or discard the conversion. Right-click can choose Keep Original or Replace Original before conversion; replacement still occurs only after all shapes were created successfully. Rotated tables are rejected because their cells cannot be reproduced safely as unrotated rectangles.

The conversion retains plain text and basic fill, line, font, alignment, orientation, and margin settings. Rich text, hyperlinks, table-specific borders, merged-cell semantics, gradients, patterns, and other advanced table formatting may not be retained. Review the generated shapes before deleting the original table.

### Copy Column Widths

**Command ID:** `BtnTableColumnWidth` · **Type:** Command · **Maturity:** Stable

Copies all column widths from the selected table for later application to a compatible table.

**How to use**

Click to copy every column width from the selected table for the current add-in session. Right-click to copy, paste, make all selected-table columns equally wide while preserving total width, or clear the saved widths. Paste skips targets with a different column count instead of resizing them partially.

### Copy Row Heights

**Command ID:** `BtnTableRowHeight` · **Type:** Command · **Maturity:** Stable

Copies all row heights from the selected table for later application to a compatible table.

**How to use**

Click to copy every row height from the selected table for the current add-in session. Right-click to copy, paste, make all selected-table rows equally high while preserving total height, or clear the saved heights. Paste skips targets with a different row count instead of resizing them partially.

### Copy/Paste Table Dimensions

**Command ID:** `BtnTableDimensions` · **Type:** Command · **Maturity:** Stable

Copies and pastes a table's column widths, row heights, or both through the feature context menu.

**How to use**

Select a source table, then choose **Copy Column Widths**, **Copy Row Heights**, or **Copy Both** from the context menu. Select one or more target tables and choose the matching paste command. A target must have the same number of relevant columns or rows; mismatched tables are skipped instead of being changed partially. Copied values remain available for the current PowerPoint add-in session only.

### Insert Column Left

**Command ID:** `BtnInsertColumnLeft` · **Type:** Command · **Maturity:** Stable

Inserts a new column to the left of the selected column.

**How to use**

Inserts one column before the leftmost selected table column. Right-click to preserve the total table width, duplicate the selected column's basic formatting into an empty column, or insert 2–50 columns. Multiple insertion can optionally preserve the total width.

### Insert Column Right

**Command ID:** `BtnInsertColumnRight` · **Type:** Command · **Maturity:** Stable

Inserts one or more columns to the right of the selected column, with optional width preservation or format duplication.

**How to use**

Inserts one column after the rightmost selected table column. The standard action grows the table. Right-click for **Insert Column Right & Preserve Width**, **Duplicate Selected Column Right – Format Only**, or **Insert Multiple Columns…**. Format-only duplication copies the selected column's width and basic cell formatting while leaving all new cells empty. Multiple insertion supports 2–50 columns and can optionally preserve the total table width.

### Insert Row Above

**Command ID:** `BtnInsertRowTop` · **Type:** Command · **Maturity:** Stable

Inserts a new row above the selected row.

**How to use**

Inserts one row before the topmost selected table row. Right-click to preserve the total table height, duplicate the selected row's basic formatting into an empty row, or insert 2–50 rows. Multiple insertion can optionally preserve the total height.

### Insert Row Below

**Command ID:** `BtnInsertRowBottom` · **Type:** Command · **Maturity:** Stable

Inserts a new row below the selected row.

**How to use**

Inserts one row after the bottommost selected table row. Right-click to preserve the total table height, duplicate the selected row's basic formatting into an empty row, or insert 2–50 rows. Multiple insertion can optionally preserve the total height.

### Insert Table Column

**Command ID:** `BtnTableColumnInsertion` · **Type:** Command · **Maturity:** Stable

Inserts a new column into the table (left by default).

**How to use**

Select one or more cells and run the command to insert a column before the leftmost selected column. Right-click to insert left or right, preserve the total table width, or duplicate the adjacent column's basic formatting into an empty column. Each successful action creates one PowerPoint Undo entry.

### Insert Table Row

**Command ID:** `BtnTableRowInsertion` · **Type:** Command · **Maturity:** Stable

Inserts a new row into the table (above by default).

**How to use**

Select one or more cells and run the command to insert a row above the topmost selected row. Right-click to insert above or below, preserve the total table height, or duplicate the adjacent row's basic formatting into an empty row. Each successful action creates one PowerPoint Undo entry.

### Move Table Rows or Columns

**Command ID:** `BtnMoveTableRowColumn` · **Type:** Command · **Maturity:** Stable

Moves, swaps, or duplicates selected table rows and columns, including contiguous blocks.

**How to use**

Select one or more cells, or place the text cursor inside a cell. Move the corresponding row or column by one step, to an edge, or to a numbered position. Contiguous selections move as one block. **Keep table design** moves formatted content while retaining destination fills, borders, dimensions, and cell layout. Other modes can move all formatting, retain only destination borders, or move plain text only. The context menu also supports swapping two selected rows or columns, duplicating a selected block, repeating the last move, and protecting the first or last row.

### Sort A-Z

**Command ID:** `BtnTableSortAZ` · **Type:** Command · **Maturity:** Stable

Sorts the selected cells ascending with culture-aware numeric ordering when every compared value is numeric.

**How to use**

Select at least two cells vertically in the sort column. The primary command sorts only those selected cell texts and keeps destination formatting in place. The **Sort Table Cells** context menu also provides descending and whole-row-text variants. Numeric values use culture-aware numeric order, text uses culture-aware case-insensitive order, and equal values remain stable. Read or write errors trigger rollback of the affected texts and successful sorts create a PowerPoint Undo entry.

### Sort Table Cells

**Command ID:** `BtnTableSort` · **Type:** Command · **Maturity:** Stable

Opens ascending and descending sorting for selected cells or complete row texts.

**How to use**

Select at least two vertical cells in the sort column, then choose A-Z or Z-A. Selected-cell modes reorder only that column's plain text. Whole-row modes reorder every cell text in the selected row range using the leftmost selected column as key; destination formatting stays in its row. Numeric values sort numerically, text uses culture-aware case-insensitive comparison, and failures attempt to restore all affected texts.

### Split Table

**Command ID:** `BtnTableSplit` · **Type:** Command · **Maturity:** Stable

Splits a selected table before the active column, or before the active row from the context menu.

**How to use**

Select a cell that should become the first cell of the second table. A normal click splits before that column; use **Split Table by Row** in the context menu to split before that row. The command builds both results before replacing the original, so an unsupported merged-cell boundary or another conversion error leaves the source table unchanged. The second table is placed 20 points to the right or below the first.

### Table Format Painter

**Command ID:** `BtnTableFormatPainter` · **Type:** Command · **Maturity:** Stable

Paints fill or basic text formatting from one selected table cell along chosen row and column directions.

**How to use**

Select exactly one source cell, then choose one or more directions. Right-click for zero-skip horizontal, vertical, or all-direction presets that copy text formatting, cell fill, or both. **Text Formatting** copies the basic font, paragraph alignment, vertical anchor, and margins; **Cell Fill** copies a visible solid fill and transparency. Borders, rich-text runs, hyperlinks, gradients, patterns, and effects are not copied. **Cells to skip between paints** controls the interval: 0 paints every cell, 1 every second cell, and so on. The palette is bound to PowerPoint and only one instance opens at a time.

### Table Style Library

**Command ID:** `BtnTableBranding` · **Type:** Command · **Maturity:** Stable

Saves reusable table styles and applies their header, body, alternating-row, border, font, and uniform-height settings.

**How to use**

Save the selected table's detected header, body, alternating-row, font, four-side border, and uniform row-height styling under a unique name. Applying a style changes formatting but keeps cell text and table structure. Saving an existing name asks before replacement, storage errors are reported, and deleting a style requires confirmation. Header detection uses PowerPoint's table setting when available and otherwise compares the first body rows. Non-uniform source row heights are not flattened into one saved value.

### Transpose Table

**Command ID:** `BtnTableTranspose` · **Type:** Command · **Maturity:** Stable

Transposes the selected table in place or creates a positioned copy while retaining its overall bounds and basic cell formatting.

**How to use**

Swaps rows and columns in the selected table. The primary action replaces the source in place. Right-click can instead keep the original and place a transposed copy to its right or below it, clamped to the slide bounds. Relative column widths become relative row heights and vice versa. Plain text, solid fills, basic font and alignment settings, margins, and the four cell borders are copied. Rich-text runs, hyperlinks, gradients, patterns, effects, and merged cells are not supported. Merged tables are rejected without modifying the source, and any creation error leaves the original table unchanged.

## Visual elements

### Add Analytical Chart Overlays

**Command ID:** `BtnChartOverlay` · **Type:** Command · **Maturity:** Stable

Calculates mathematical offsets over standard charts, including a chart inside a selected group

**How to use**

Select exactly one native PowerPoint chart or one group containing exactly one chart. Right-click preselects Delta, CAGR, Growth, Average Level Line, or Highlight Outliers; the dialog still lets you review point ranges and styling before insertion. The chart remains grouped, while the generated overlay is added to the slide as an independent grouped annotation. Overlays are mathematical PowerPoint shapes and support stacked-column total aggregation.

### Add Corner Sticker

**Command ID:** `BtnAddSticker` · **Type:** Command · **Maturity:** Stable

Adds a small colored sticker (rounded rectangle) to the top-right corner of the slide.

**How to use**

Inserts a compact rounded status sticker in the top-right corner of the current slide. The primary action uses the configured default text and font. Right-click inserts ready-made Draft, Review, Approved, or Confidential stickers.

### Checkbox

**Command ID:** `BtnCheckbox` · **Type:** Command · **Maturity:** Stable

Inserts or updates a framed or unframed Tick, Cross, or Empty checkbox. Default is a framed Tick.

**How to use**

Inserts a framed Tick checkbox at the center of the slide, or updates one selected tagged checkbox. Right-click to insert or apply Tick, Cross, and Empty presets with or without a frame. Existing tagged checkbox shapes retain their position and size.

### Copy Text Margins

**Command ID:** `BtnLearnMargin` · **Type:** Command · **Maturity:** Stable

Copies the left, right, top, and bottom internal text margins from one shape for use with Paste Text Margins.

**How to use**

Select exactly one shape with a text frame to copy its left, right, top, and bottom internal text margins. These are the distances between the text and the shape edges, not gaps between shapes. The values are stored for the current add-in session without changing the Windows clipboard. Select target shapes and use **Paste Text Margins** to apply them. Right-click to copy again or clear the copied values.

### Generate Shape Series

**Command ID:** `BtnSeriesGenerator` · **Type:** Toggle · **Maturity:** Stable

Creates or updates a horizontal, vertical, or matrix series of numbered, lettered, dated, or quarter-labeled shapes.

**How to use**

Optionally select a template shape first, then create up to 1,000 number, letter, day, month, year, or quarter labels. Right-click opens common 1–5 number, A–E letter, vertical, or 3×3 matrix presets while keeping the preview editable. Choose a horizontal, vertical, or matrix layout and optionally fit it to the slide or a second selected boundary shape. Generated shapes retain series metadata, and a complete generation or update creates one PowerPoint Undo entry.

### Harvey Ball

**Command ID:** `BtnHarveyBall` · **Type:** Command · **Maturity:** Stable

Cycles through Harvey Ball states (0%, 25%, 50%, 75%, 100%) to represent progress or qualitative data.

**How to use**

Inserts or updates a circular progress indicator (Harvey Ball).
**Usage:**
* Left-click to cycle through 0 -> 25 -> 50 -> 75 -> 100% fill.
**Right-Click Options:**
* Jump directly to a specific fill level (0%, 25%, 50%, 75%, or 100%).

### Map Builder

**Command ID:** `BtnMapWizard` · **Type:** Command · **Maturity:** Stable

Inserts editable world, continent, country, and administrative-region maps, and colors their regions from pasted tabular data.

**How to use**

Choose **World**, **Continent**, or **Country or territory**, then choose the available detail level. Right-click starts at World, Continent, Country, the selected tagged map, or the most recent configuration. Administrative levels are enabled only where bundled data is available, and every inserted region is an editable native PowerPoint shape.

Choose whether remote regions stay in place, move into compact insets, or are excluded. Use **Mercator** or **Natural proportions**, restore a recent selection, and color shapes by pasting region names/IDs and values from Excel.

### Numbered Marker

**Command ID:** `BtnNumeration` · **Type:** Command · **Maturity:** Stable

Inserts one editable numbered marker on the current slide.

**How to use**

Insert one circular numbered marker on the current slide. Right-click for direct number 1 through 5 presets. Select the marker to edit its number, size, color, opacity, or shape from the contextual toolbar. The marker stores its settings in openEfficiency metadata and can be changed without overwriting text in other shapes.

### Paste Text Margins

**Command ID:** `BtnSetMargin` · **Type:** Command · **Maturity:** Stable

Pastes the internal text margins saved by Copy Text Margins onto the selected shapes with text frames.

**How to use**

Use **Copy Text Margins** on one source shape first, then select the target shapes. This command applies all four copied internal text margins and skips shapes without text frames. Values come from the current add-in session, not the Windows clipboard. Right-click to paste all margins, only left and right, only top and bottom, or reset every margin to zero. Reset does not require copied values. Each action creates one Undo entry.

### Progress Series

**Command ID:** `BtnProgressSeries` · **Type:** Command · **Maturity:** Stable

Inserts a progress series visualization with configurable shape type and layout.

**How to use**

Inserts an editable numbered or bullet progress series. Right-click for common three-step, five-step, horizontal, vertical, and bullet presets; selecting an existing series updates it in place. If you manually recolor the active segment in PowerPoint, opening the contextual toolbar reads that visible fill without ungrouping the series and retains it until you choose another color.

### QR Code

**Command ID:** `BtnQRCode` · **Type:** Toggle · **Maturity:** Stable

Downloads a QR code for entered text and inserts it as a PowerPoint image on the current slide.

**How to use**

Enter a URL or text string. openEfficiency sends that content to the external QR Server service, downloads the generated PNG, and inserts it at the center of the current slide. The temporary image is always removed after the attempt. Generation requires an internet connection and creates one PowerPoint Undo entry.

### Repeat Shape

**Command ID:** `BtnRepeatShape` · **Type:** Command · **Maturity:** Stable

Clones the selected shape multiple times to create long sequences or grids. Default direction is Right.

**How to use**

Duplicates the selected shape a configurable number of times with automatic offset spacing, instantly generating a row or column of identical shapes.

### Split Shape into Grid

**Command ID:** `BtnSplitShape` · **Type:** Command · **Maturity:** Stable

Splits the selected shape into a grid of smaller shapes (default 2x2 with no spacing).

**How to use**

Divides one selected shape into an N × M grid of equally-sized non-overlapping shapes. Right-click for direct 1×2, 2×1, 2×2, 3×2, or 3×3 presets with zero spacing. Direct presets replace the source only after every grid shape has been created successfully.

### Star Rating

**Command ID:** `BtnStarRating` · **Type:** Command · **Maturity:** Stable

Inserts a sequence of stars representing a rating from 1 to 5. Default is 5 stars, 3 rating.

**How to use**

Inserts a customizable star rating indicator (1 to 5 stars, with half-star support filled from left to right) as a native PowerPoint vector group. Right-click for direct 1-of-5 through 5-of-5 presets; a selected rating is updated in place.

### Status Sticker

**Command ID:** `BtnIllustrativeSticker` · **Type:** Command · **Maturity:** Stable

Adds and manages consulting status stickers, including workflow, confidentiality and audience labels across one or many slides.

**How to use**

Click to insert the default status label. Right-click to choose presets or a custom label, combine workflow, confidentiality and audience states, update an existing sticker, apply it to selected slides or the entire presentation, change its corner position, use the presentation theme color, add date or author metadata, remove stickers, or manage presets in Settings. Existing tagged stickers are updated instead of duplicated by bulk actions.

### Thermometer

**Command ID:** `BtnThermometer` · **Type:** Command · **Maturity:** Stable

Inserts a thermometer visualization for project progress. The default fill level is 50%.

**How to use**

Inserts or updates a vertical thermometer-style progress bar.

**Usage:**
* With no thermometer selected, left-click to insert one at 50%.
* Select an existing thermometer and left-click to increase its fill by 25%; after 100% it wraps to 0%.

**Right-click options:**
* Set the fill level exactly (0%, 25%, 50%, 75%, or 100%). Existing thermometers keep their current color.

### Traffic Light

**Command ID:** `BtnTrafficLight` · **Type:** Command · **Maturity:** Stable

Inserts a traffic light infographic (red/yellow/green states). Default is Green state.

**How to use**

Click to insert a green traffic light as a native PowerPoint shape group. When an existing openEfficiency traffic light is selected, clicking sets it to green while preserving its orientation, frame, and frame weight. Right-click the Sidebar button to insert or set Red, Yellow, Green, or Off directly. Select a traffic light to use the contextual controls for color, orientation, and frame settings. State and appearance settings persist when the presentation is saved, reopened, or the shape is copied.

## Storyline and review

### Agenda Creator

**Command ID:** `BtnAgendaWizard` · **Type:** Command · **Maturity:** Stable

Creates or synchronizes a timed agenda and its tagged dividers, sections, starter slides, and optional backup structure from a reusable local layout.

**How to use**

Define agenda topics, hierarchy, durations, speakers, visibility, start time, and target end time, then choose a layout. The built-in Standard Agenda is selected automatically for a clean profile, so an agenda can be created without first designing a custom layout. Create or synchronize a main agenda, chapter divider slides, PowerPoint sections, optional starter slides, highlighted current topics, and a tagged backup slide. Existing tagged agendas are updated in place while chapter content is retained by stable agenda IDs. Use the context menu to prepare a full deck structure, derive an agenda from sections and slide titles, validate the agenda against the deck, or open Agenda Layout Designer. The wizard can import external AgendaCreator.xml and legacy AgendaWizard.xml layout definitions without replacing existing local layouts. Generation uses a PowerPoint Undo transaction and restores the previous agenda when a reported generation error occurs. Presentations or affected structures protected by think-cell are left unchanged. Saved agenda definitions are stored locally; saving an existing name asks before replacement.

### Agenda Layout Designer

**Command ID:** `BtnSaveAgendaLayout` · **Type:** Command · **Maturity:** Stable

Captures formatting and placement from selected shapes as a reusable local Agenda Creator layout.

**How to use**

Assign shapes from the active slide to agenda roles, validate their dimensions and spacing against the actual slide size, and save the result as a local layout. Saving an existing custom name requires confirmation. The built-in Standard Agenda can be inspected and reused but cannot be overwritten or deleted. A clean local template copy is generated without the assigned marker shapes; the source slide is not changed. Custom layout configuration and templates are stored together under the local AgendaLayouts folder. This command saves layout formatting, not agenda topics, durations, or speakers.

### Anonymize Selected Slide Text

**Command ID:** `BtnAnonymize` · **Type:** Command · **Maturity:** Stable

Replaces visible text on selected slides with placeholder words while retaining text-frame structure and character formatting where PowerPoint permits. Dynamic slide numbers are preserved.

**How to use**

Replaces visible text on the current slide, selected slides, or the entire presentation. Right-click selects the scope or previews the affected word count without changing content. The command processes ordinary text boxes, grouped shapes, and table cells while retaining whitespace, punctuation, letter case, and character formatting where PowerPoint permits. It preserves dynamic slide-number placeholders and asks for confirmation before replacement.

This command does not remove speaker notes, comments, document properties, alternative text, chart data, or embedded-file metadata. Remove those items separately when required.

### Consolidate Slide Masters

**Command ID:** `BtnConsolidateMasters` · **Type:** Command · **Maturity:** Stable

Backs up the presentation, remaps slides to the primary master, flags migrated slides for review, and removes redundant masters.

**How to use**

Creates a timestamped backup copy before changing the open presentation. Slides that use another master are mapped to an equally named layout on the primary master where possible, with Blank or the first layout as a fallback. Migrated slides receive a tagged review marker. If slide mapping fails, completed mappings are rolled back and no masters are deleted. If master deletion fails, the command reports the incomplete result and the backup path.

### Create Slide Guide Grid

**Command ID:** `BtnSlideGuidelines` · **Type:** Command · **Maturity:** Stable

Creates responsive tagged guide shapes on the current slide and replaces only guides owned by openEfficiency.

**How to use**

Creates responsive line-shape guides for the current slide size. A normal click applies a two-column layout; right-click offers three, four, five, 1:2, and rule-of-thirds layouts plus explicit removal. Applying a layout replaces only shapes tagged as openEfficiency guidelines, preserving similarly named user shapes. These are slide shapes rather than PowerPoint's native guides.

### Export Presentation

**Command ID:** `BtnExportWizard` · **Type:** Toggle · **Maturity:** Stable

Exports all or selected slides as PPTX, PDF, both formats, or editable HTML, with optional cleanup and delivery actions.

**How to use**

Choose the entire presentation or selected slides, optionally exclude hidden slides, and export to PPTX, PDF, both formats, or editable HTML. Right-click opens the export window with a common format and scope combination already selected; export never starts until you confirm it. Cleanup applies only to exported copies; the open presentation is changed only when **Also apply these properties to the open presentation** is explicitly enabled. Existing output files are listed for confirmation before replacement. Clipboard and e-mail delivery use temporary files, while disk export opens the destination folder when complete.

### Master Covers & Overlays

**Command ID:** `BtnFlightMode` · **Type:** Toggle · **Maturity:** Stable

Temporarily covers master-slide images or adds a light or heavy white overlay across every slide master.

**How to use**

Temporarily adds tagged, removable shapes to every slide master. Toggle defaults to Cover Logos, which places white rectangles over master pictures. The context menu also offers Light Overlay (95% of the slide remains visible), Heavy Overlay (85% remains visible), and Disable Master Covers & Overlays. Switching modes first removes the previous Master Covers & Overlays shapes; disabling restores the original masters.

### Move Slides to Backup

**Command ID:** `BtnMoveToBackup` · **Type:** Command · **Maturity:** Stable

Moves selected slides, in their existing order, to the end of a presentation that contains a designated Backup slide.

**How to use**

Moves the selected slides to the end of the presentation, behind the Backup slide created by Agenda Creator. Right-click can move, copy, or locate the designated Backup slide. This feature never creates the Backup slide. Move and copy stop without changes when no designated Backup slide exists, exclude the Backup slide itself, require confirmation, and start a PowerPoint Undo entry where supported.

### New Slide

**Command ID:** `BtnNewSlide` · **Type:** Command · **Maturity:** Stable

Opens PowerPoint's native New Slide gallery without inserting a slide until a layout is chosen.

**How to use**

Open PowerPoint's native New Slide gallery and choose a layout from the available slide masters. Right-click to open the gallery, insert a blank slide after the current slide, or duplicate the current slide. Creation and Undo behavior remain native to PowerPoint where applicable.

### Presentation Cleaner

**Command ID:** `BtnCleaner` · **Type:** Command · **Maturity:** Stable

Reviews, sanitizes, or removes presentation artifacts with explicit scope and confirmation for destructive actions.

**How to use**

Opens cleanup and review actions for hidden slides, animations, transitions, notes, comments, unused masters and layouts, sticky notes, footers, repeated spaces, accessibility text, shape locks, ink, external data links, invisible or off-slide objects, document information, unprotected custom XML, and visible numbers. Destructive actions show their scope and require confirmation; PowerPoint Undo is started where supported. Shape cleanup and XML cleanup preserve oPE-owned objects and recognizable Efficient Elements or think-cell content. Convert Slides to Images creates a separate presentation and leaves the source unchanged. Anonymization and number sanitization have separate confirmations.

### Private Review Notes

**Command ID:** `BtnSlideNotes` · **Type:** Toggle · **Maturity:** Stable

Stores multiple private review notes per slide, with Open or Done status and optional shape anchoring.

**How to use**

Create multiple private review notes on the current slide, optionally attach them to a selected shape, and manage them as **Open** or **Done**. Right-click the feature to open the presentation-wide overview or delete notes with confirmation. Notes are stored as presentation metadata and stay out of PowerPoint's review UI and slideshow, but they are not encrypted and are not a security boundary.

### Sticky Note

**Command ID:** `BtnStickyNote` · **Type:** Command · **Maturity:** Stable

Inserts a sticky note at the top-right of the slide. Style is configurable in Settings.

**How to use**

Inserts a styled sticky-note annotation at the top-right of the current slide. Theme, Consulting, and Compact Banner each keep their own background and font colors in Settings; Compact Banner also uses its characteristic offset shadow block. Default text and font are shared. New notes are marked with an openEfficiency tag so Sticky Note Manager can identify them safely.

### Sticky Note Manager

**Command ID:** `BtnStickyNoteManager` · **Type:** Command · **Maturity:** Stable

Moves, restores, deletes, or creates sticky notes from PowerPoint comments.

**How to use**

A normal click moves notes off the current slide without deleting them. The context menu provides restore and confirmed delete actions for the current or selected slides, copies PowerPoint comments to tagged sticky notes while retaining the originals, and marks selected shapes as sticky notes.

### Storyline

**Command ID:** `BtnStoryline` · **Type:** Command · **Maturity:** Stable

Reviews, filters, edits, and exports section-aware slide headlines with consulting-focused quality checks.

**How to use**

Open a section-aware narrative view of the presentation. Search or filter headlines, navigate through stable slide IDs, edit detected title shapes, assign review status, and use the context menu for slide actions. The quality check flags missing, duplicate, generic, overly long, question-only, and weak topic headlines and provides local action-title suggestions without sending presentation content to an external service. Copy or export the complete deck, the current section, headlines only, or issues only.

### Style Check

**Command ID:** `BtnStyleCheck` · **Type:** Command · **Maturity:** Stable

Checks the current slide, selected slides, or the entire presentation for style and consulting-deck consistency issues.

**How to use**

Choose **Current Slide**, **Selected Slides**, or **Entire Presentation** before scanning. Style Check identifies off-brand fonts and colors, spacing and text-overflow issues, distorted images, overlapping text, missing or overly long titles, missing source lines on chart or table slides, footer and pagination inconsistencies, and common wording defects. Use **Jump** to select the affected object, **Ignore Rule** to exempt only that rule for the object or slide, and **Fix Safe Issues...** to preview and apply low-risk color and wording corrections. Right-click the Sidebar button to open a specific scope directly.

## Data and automation

### Build Presentation from Answers

**Command ID:** `BtnDocumentAutomation` · **Type:** Command · **Maturity:** Stable

Builds a versioned question schema in the presentation and creates automated copies with text replacements and simple slide conditions.

**How to use**

Use **Design Schema** to add, edit, remove, reorder, validate, and save Text, Yes/No, and Single Choice questions. Place `{{Key}}` tokens in ordinary shapes, nested groups, or table cells; surrounding character formatting is preserved. Add at most one simple Equals or NotEquals condition per slide, based on a Yes/No or Single Choice answer. Use **Run Automation** to review detected placeholders, missing values, slide inclusion decisions, and validation errors. Generation always creates and edits a new `.pptx` copy; the template remains unchanged. Corrupt stored schema parts fall back to the newest valid version, while unsupported newer schema versions are rejected. Concurrent runs may use different targets, but a second run to the same target fails safely. Schemas are limited to 200 questions, 1,000 slide conditions, 100 options per question, 64-character keys, 500-character prompts, 200-character options, and 10,000-character answers. REST, external data sources, image questions, nested conditions, and formulas are not supported in this MVP.

### Excel Link Manager

**Command ID:** `BtnExcelLinkManager` · **Type:** Command · **Maturity:** Stable

Lists linked Excel objects and native linked tables, with refresh, source navigation, relinking, and unlinking actions.

**How to use**

Open the manager to inspect Excel links across the active presentation. Select one or more entries to refresh them or jump to their slide. Native linked tables also support Open Source and Relink Source. Unlink is destructive and keeps the current PowerPoint content while removing its source connection. Embedded charts are intentionally excluded because they have no external Excel link.

### Gantt & Timeline Editor

**Command ID:** `BtnGanttChart` · **Type:** Toggle · **Maturity:** Stable

Creates rule-based Gantt charts or flexible project timelines directly in PowerPoint.

**How to use**

Open the editor and click **Create New Chart** to insert a new chart with one editable starter task. Use **Update Active Chart** to rebuild the chart currently loaded in the editor. Use the editor in two complementary ways: build a rule-based Gantt chart with task hierarchies, consistent working-day durations, network-based critical-path calculation, and jointly evaluated FS/SS/FF predecessor constraints, or create a flexible Think-cell-style timeline with freely arranged lanes, events, sub-milestones, and sub-bars. You can switch between both modes; re-enabling Gantt rules recalculates summaries. Select an existing openEfficiency chart before opening the editor to continue editing that chart; multiple independent charts can coexist on one slide. Summary tasks move together with their descendants. New sibling rows are inserted after the complete subtree, deleting a summary retains and promotes its descendants, and Indent is available only when a valid parent row exists above. Each structural edit is restored as one complete Undo step. Long task names use 0.1 cm internal margins, wrap inside the fixed-width name column, and grow their lanes automatically. Select a regular task bar to show handles at both ends; drag the left handle to change its start or the right handle to change its end. You can also edit task names or schedules directly on the slide and use Sync From Slide to bring those changes back into the editor. Excel sync validates dates and preserves existing PowerPoint-only styling when rows can be matched by WBS or task name. Right-click chart elements for row, milestone, event-line, indentation, and styling actions.

### Import HTML Slides

**Command ID:** `BtnHtmlSlideImport` · **Type:** Command · **Maturity:** Stable

Imports supported AI-generated presentation HTML files and rebuilds them as editable PowerPoint slides.

**How to use**

Choose a supported AI-generated HTML presentation, review the detected slides, and select which ones to import. Right-click presets source styling, presentation theme, background-only, or local-resources-only behavior; every mode still shows the detected slides before import. Recognized sections, cards, lists, callouts, and process flows become editable PowerPoint shapes. JavaScript is never executed. Local Resources Only never requests remote files. For stability, HTML is limited to 5 MiB and stylesheet downloads are bounded by size, count, redirects, and time.

### Insert Excel Value

**Command ID:** `BtnInsertExcelValue` · **Type:** Command · **Maturity:** Stable

Inserts the active Excel cell as a refreshable linked text box or inline value in PowerPoint.

**How to use**

Keep Excel open and select one cell. In PowerPoint, place the text caret inside a text box to insert the linked value inline; otherwise the command creates a linked text box at the pointer position. Right-click explicitly chooses inline, linked text-box, or unlinked insertion, and can refresh or unlink the selected value. Select a linked result to adjust decimals, thousands separators, and percent formatting. Unlink keeps the displayed text and removes only its Excel connection.

### Insert Linked Text Label

**Command ID:** `BtnConnectedTextLabel` · **Type:** Command · **Maturity:** Stable

Publishes selected PowerPoint shape text as a refreshable label in an open Word document or OneNote page.

**How to use**

Select exactly one PowerPoint shape containing text. The primary action sends it to the active Word document; if Word is unavailable, it tries OneNote. Right-click to choose Word or OneNote explicitly, refresh every connected label whose PowerPoint source is in the active presentation, or disconnect the selected source. Word labels are stored as tagged content controls. OneNote labels are stored as tagged page elements. OneNote sends stop without writing when the visible current page cannot be identified or its version changes during the update. Refresh updates open Word documents and remembered OneNote pages; it never opens or saves Office files without an explicit send action. The link stores text only and does not copy formatting, images, tables, notes, or hyperlinks.

### Insert Slide Snapshots

**Command ID:** `BtnSlidePaste` · **Type:** Command · **Maturity:** Stable

Selects slides from the current deck and inserts stable image references on the active slide in an automatically sized grid.

**How to use**

Opens thumbnails for the presentation that was active when the window loaded. Right-click presets the hyperlink and shadow combination. Select source slides, then confirm Insert Selected to place PNG snapshots on the active slide. The command aborts if you switch presentations, resolves reordered source slides by Slide ID, groups the insertion into one PowerPoint Undo entry, tags each image with its source Slide ID, and reports skipped exports. Temporary thumbnail and export files are removed when the window closes.

### Inspect Fonts & Colors

**Command ID:** `BtnPropertyExtraction` · **Type:** Command · **Maturity:** Stable

Displays unique fonts, sizes, text colors, fill colors, and line colors in the selection or current slide.

**How to use**

Inspects selected shapes or the current slide. Right-click chooses the scope and can copy the full report or a deduplicated hexadecimal color list directly. Reports include grouped shapes, mixed-format text runs, and table-cell text, plus unique fonts, sizes, text colors, fills, and lines. Picture fills are skipped because they have no single fill color.

### Rebuild Slide from PDF/Image

**Command ID:** `BtnLocalSlideRebuild` · **Type:** Command · **Maturity:** Stable

Uses Windows OCR and PDF rendering on this device to create a local reference slide with editable text, background, and accent shapes.

**How to use**

Choose an image or PDF. Right-click can restrict the picker to images or PDFs, or request selected PDF pages such as 1-3,5. The source remains on your computer: Windows renders PDF pages locally and Windows OCR detects text locally. A progress window shows the current rendering and OCR step.

The feature creates a native background color and long horizontal accent lines, then adds the original page as a reference background and overlays editable PowerPoint text boxes at the detected positions. It automatically masks detected source text when the surrounding background is a single color, so edits do not leave duplicate text behind. Complex graphics, diagrams, text formatting, photos, and gradients remain part of the reference image. For best results, use high-resolution source files and install the Windows language pack used by the source text.

### Refresh Linked Excel Charts

**Command ID:** `BtnUpdateExcelCharts` · **Type:** Command · **Maturity:** Stable

Updates externally linked Excel charts on the current slide or across the presentation.

**How to use**

Click the feature and choose whether to update the current slide or the complete presentation. Only externally linked Excel charts are refreshed; embedded charts are left unchanged. A keyboard shortcut updates the current slide directly.

### Refresh Linked Table

**Command ID:** `BtnRefreshLinkedTable` · **Type:** Command · **Maturity:** Stable

Reloads the linked Excel source of the selected table and writes the latest values into the table cells. The table must have been created via 'Link as Table' in Excel first.

**How to use**

Select exactly one table previously created through Excel's Link as Table command, then refresh it to load the latest source range. Value-only links preserve PowerPoint formatting; formatted links also apply supported Excel formatting. If Auto-Resize is enabled, the table grows or shrinks to match the source range. Missing files and invalid ranges are reported without removing the existing table content.

### Tag Inspector

**Command ID:** `BtnTagInspector` · **Type:** Command · **Maturity:** Stable

Allows viewing, adding, and deleting hidden metadata tags on the current slide, the presentation, or selected shapes.

**How to use**

Inspect hidden metadata stored on the presentation, current slide, or selected shapes. Right-click to open directly at any of those three scopes. Mixed values across selected shapes are shown explicitly; edits apply to every selected shape. Removing openEfficiency system tags can disable related features, so deletion is confirmed. Sticky-note conversion copies content and retains the source.

## Libraries

### Asset Library

**Command ID:** `BtnAssetManager` · **Type:** Toggle · **Maturity:** Stable

Opens a searchable local library for reusable slides, shapes, images, folders, tags, and favorites.

**How to use**

Use Slides, Images, Shapes, and My Assets to browse or search reusable local content. Right-click opens a specific library section or starts Save Current Selection directly. Save the current slide or shape selection into a chosen Favorites folder, optionally with tags and automatic tagging. PowerPoint export and import work remains on PowerPoint's UI thread; generated files use unique names, and rename or deletion operations are restricted to the Asset Library root. Permanent deletion requires confirmation in the relevant library tab.

### Country Flags

**Command ID:** `BtnCountryFlags` · **Type:** Command · **Maturity:** Stable

Searches the bundled ISO 3166-1 flag catalog and inserts the selected flag as vector SVG content.

**How to use**

Search by country name, alpha-2, alpha-3 or a common alias. Recent flags appear first. Enter or double-click inserts the bundled SVG at its native aspect ratio in one PowerPoint Undo step; normal use is fully offline.

### Icon Library

**Command ID:** `BtnIconBrowser` · **Type:** Command · **Maturity:** Stable

Browse licensed icon packs, filter by style and theme, recolor icons, and insert them as SVG.

**How to use**

Browse licensed icon packs, filter by style and theme, recolor icons, and insert them as SVG images or editable PowerPoint shapes into the active slide. Right-click opens focused search, Favorites, or a predefined insert mode.

### Illustration Library

**Command ID:** `BtnIllustrationLibrary` · **Type:** Command · **Maturity:** Stable

Pick a bundled or imported illustration, recolour it, switch its parts off or on, and insert it into the slide - or open the evaluated external sources and import your own SVG artwork.

**How to use**

The Illustration Library has two views. Figures: every bundled and imported illustration with its palette (each colour can be replaced) and its named parts (each part can be left out); insert it as a scalable SVG picture or as converted editable shapes, and the inserted shape carries its source in the shape tags. Sources & import: the bundled CC0 collections plus every evaluated external illustration source with its license, attribution duty and technical notes. Open a source in your browser, download the artwork you are licensed to use, and import the SVG files: the import screens each file for scripts, embedded images and external references, refuses files it cannot recolour, and records the source and license in the user library below %LOCALAPPDATA%\oPenEfficiency\Illustrations. Only the bundled CC0 collections may be redistributed with the add-in; licensing of imported artwork stays your own responsibility.

### Template Library

**Command ID:** `BtnTemplateManager` · **Type:** Command · **Maturity:** Stable

Find and open PowerPoint templates from default locations, watched folders, and explicitly added files.

**How to use**

Search default Windows and Office template locations, add watched folders, or register individual POTX, POT, PPTX, and PPTM files. Right-click opens search, file registration, watched-folder management, or hidden-template restoration directly. Open creates an untitled presentation so the source template remains unchanged. Hide removes an entry from the library view without deleting its file. Permanent deletion is available only for files explicitly added to the library and always requires confirmation.

## Settings and help

### Command Search

**Command ID:** `BtnSearchBar` · **Type:** Widget · **Maturity:** Stable

Find and execute any openEfficiency command by canonical or legacy name.

**How to use**

Type a command name, legacy name, or stable ID to filter the Sidebar. Press Enter to run the best match, or Escape to clear the search and return to the configured layout. Right-click the widget to focus the search field or clear the current query.

### Feature Explorer

**Command ID:** `BtnExploreFeatures` · **Type:** Command · **Maturity:** Stable

The Feature Explorer provides a searchable index of all commands available in openEfficiency.

**How to use**

Search the openEfficiency command catalog by canonical name, legacy name, description, or stable ID, then review practical usage guidance and the current shortcut. Press Ctrl+F to focus search or Escape to close. Right-click to focus all-feature search or open common catalog categories directly.

### Performance Mode

**Command ID:** `BtnPerformanceMode` · **Type:** Toggle · **Maturity:** Stable

Temporarily pauses only the external COM add-ins explicitly selected by the user and restores only those changed in this session.

**How to use**

Right-click to choose detected external PowerPoint COM add-ins. Nothing is selected by default. Enabling Performance Mode pauses selected active add-ins through their supported Connect property; disabling or shutting down restores only add-ins this session paused.

### Quick Context Menu

**Command ID:** `BtnQuickContextMenu` · **Type:** Toggle · **Maturity:** Stable

Shows an opt-in, configurable set of openEfficiency commands beside PowerPoint's native shape context menu.

**How to use**

Click to enable or disable the Quick Context Menu. Right-click this Sidebar command to choose and order up to 12 commands. PowerPoint's native context menu remains available.

### Settings

**Command ID:** `BtnSettings` · **Type:** Command · **Maturity:** Stable

Configures Sidebar layout, appearance, sticky notes, translations, caching, local diagnostics, shortcuts, and portable configuration files.

**How to use**

Edit the Sidebar layout and appearance, sticky-note defaults, translation providers, cache folders, and shortcuts. Changes are applied only after Save & Apply succeeds. Translation API keys are stored locally using Windows user-scoped protection; existing plaintext settings are migrated automatically. Layout and configuration import skip unknown feature IDs. Portable exports omit API keys by default and require explicit confirmation before including them as plaintext. Under General, Create Diagnostic Bundle writes a local ZIP with runtime details, configured feature IDs, sanitized configuration metadata, and up to 512 KiB of the recent sanitized error log. It never uploads data and excludes presentation content, credential values, full paths, and machine/user identity; review the ZIP before sharing it. Restore All Defaults resets layout and appearance but intentionally leaves shortcuts unchanged. Right-click to open any settings section directly.

## Labs

### 2048

**Command ID:** `Btn2048` · **Type:** Command · **Maturity:** Experimental

Opens a compact 2048 puzzle game.

**How to use**

Use the arrow keys to combine tiles and reach 2048. Right-click to continue, start a new game, or close the game window.

### Chess

**Command ID:** `BtnChess` · **Type:** Command · **Maturity:** Experimental

Opens a floating chess board with a local computer opponent.

**How to use**

Play a complete game of chess against a built-in computer opponent. The game is independent from the current presentation. Right-click starts a new Easy, Medium, or Hard game or closes the board.

### Minesweeper

**Command ID:** `BtnMinesweeper` · **Type:** Command · **Maturity:** Experimental

Opens a compact Minesweeper game.

**How to use**

Clear the field without opening a mine. Right-click cells to mark suspected mines; right-click the Sidebar command to continue, start a new game, or close the game window.

### Random Winner Picker

**Command ID:** `BtnWinnerPicker` · **Type:** Command · **Maturity:** Experimental

Opens the random winner picker for animated selection from slide objects.

**How to use**

Animates a random selection from selected slide objects, useful for team engagement exercises and random presenter selection. Right-click picks one, two, or three winners immediately or clears the generated winner markers.
