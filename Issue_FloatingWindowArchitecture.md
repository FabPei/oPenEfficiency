## Problem Statement

The current implementation of features that spawn floating toolbars upon shape selection (e.g., Star Rating, Progress Series, Checkboxes, Numeration) has outgrown its original design and is no longer scalable.

### Key Pain Points:

1. **The "God Class" Bottleneck (`ThisAddIn.cs`):** 
   Currently, `ThisAddIn.cs` listens to `Application_WindowSelectionChange`. Inside `ProcessSelectionChange`, there is a massive `if... else if... else if` block hardcoding every single tag prefix (`"oPE_HarveyBall"`, `"oPE_StarRating"`, etc.). Adding a new smart shape feature requires modifying the core Add-in class.
2. **Manual Instance Tracking:** 
   `ThisAddIn.cs` manually declares a private variable for *each* toolbar (e.g., `_ratingToolbar`, `_progressSeriesToolbar`). It manually checks for nulls, closes old instances, creates new ones, and hooks up the `Closed` events.
3. **Inconsistent Positioning:** 
   Each toolbar is spawned manually in its own `Show...Toolbar()` method, leading to fragmented positioning logic (some use mouse coordinates, others calculate bounding boxes).
4. **Duplicated WPF Boilerplate & Lifecycle Bugs:** 
   Every floating toolbar inherits directly from `Window`. They each duplicate the `Window_Deactivated` logic to auto-close and implement manual hacks (like `if (ColorPopup.IsOpen) return;`) to prevent the window from closing when a child dropdown menu takes focus.

---

## Proposed Streamlined Architecture

To make this infinitely scalable and bug-free, we should introduce a standardized pipeline:

### A. A Shared Base Class: `SmartShapeToolbar : Window`
Instead of inheriting from `Window`, all context popups will inherit from a new `SmartShapeToolbar` base class.
*   **What it does:** Automatically handles `WindowStyle="None"`, `Topmost="True"`, and `AllowsTransparency="True"`.
*   **Deactivation:** Centralizes the `Window_Deactivated` logic. It will include a standard property like `IsSubDialogOpen` that child classes can set to `true` when they open a color picker, eliminating duplicated hacky checks.

### B. A Centralized `FloatingToolbarService`
Move all tracking out of `ThisAddIn.cs`. 
*   **What it does:** This service holds exactly *one* reference: `Window _activeToolbar`. 
*   If a new toolbar needs to spawn, the service automatically closes `_activeToolbar` and replaces it. This entirely removes the need to track `_ratingToolbar`, `_numerationToolbar`, etc. individually.

### C. Registration / Auto-Discovery Pattern
Just like how standard buttons use `[FeatureMetadata]`, we should map these interactive shapes.
*   Create a simple dictionary or interface mapping that connects a `Tag Prefix` (e.g., `"oPE_StarRating"`) to a specific `Window` type. 
*   When a selection changes, `ThisAddIn.cs` just asks the `FloatingToolbarService`: *"I found a shape with tag X. Do you have a toolbar for this?"* The service instantiates it via reflection, positions it consistently, and displays it.

---

## Proposed Implementation Phases

1. Create the `FloatingToolbarService` and the `SmartShapeToolbar` base class.
2. Refactor `ThisAddIn.cs` to delegate selection logic to the new Service.
3. Update the existing toolbars (Rating, Progress Series, Numeration, etc.) to inherit from the new base class and remove their duplicated lifecycle boilerplate.