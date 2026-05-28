# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).


## [2026.05.28] - 2026-05-28

### Added
- **Color Overlay:** New shapes feature that creates a compound shape from two selected objects. The bottom shape retains its color where it does not overlap, but takes the color of the top shape where they overlap.

### Changed
- **Swap Positions:** Right-click context menu options for the rotation anchor are now toggleable and persist as the default behavior for subsequent clicks on the main button.

## [2026.05.26] - 2026-05-26

### Changed
- **UX:** Reduced the appearance delay for detailed feature descriptions in sidebar tooltips from 1.5s to 0.5s for faster information access.

## [2026.05.25] - 2026-05-25

### Added
- **Multi-Swap:** New feature to swap multiple shapes based on selection order. Supports Top-Down, Bottom-Up, Left-Right, and Right-Left directions via right-click context menu.
- **Glass Hide:** New visual feature to overlay shapes with semi-transparent "glass" rectangles. Supports "Single Bounding Box" or "Individual Shapes" modes via right-click context menu.

## [2026.4.25] - 2026-05-20

### Added
- Clone Selection: Added new tool to clone objects perfectly adjacent in all 4 directions.

### Fixed
- Split Shape: Fixed bug where splitting labels caused an error.
- Magic Resizer: Enabled usage for single grouped shapes.
- Export Wizard: Single slide PPTX export now correctly saves only the selected slides.
- Match Size: Fixed Master option not being respected.
- Smart Eyedropper: Automatically detects and applies color to Fill, Line, or Text.

## [2026.05.11.1000] - 2026-05-11

### Added
- **Tag Inspector Enhancement:** Added bidirectional conversion between Sticky Notes and Slide Tags via right-click context menu.
- **Efficient Elements Support:** Asset Library now detects Efficient Elements installations in the `LocalAppData` folder (standard per-user install path).
- **Automated Releases:** Improved GitHub Actions to automatically generate changelogs from commit history and package ClickOnce installers for every release.

### Changed
- **Feature Metadata:** Refactored multiple text-related features (Align Text, Change Case, Character Spacing) to include rich descriptions and detailed help text.
- **UI:** Standardized context menu icons in the Tag Inspector for a more modern look.

## [2026.05.06.1420] - 2026-05-06

### Added
- **Command Palette Lite:** A new searchable command bar widget for the sidebar, allowing instant access to any feature.
- **Enhanced Snapping:** Alignment tools now support single-shape centering on slide guides.
- **Public Release Preparation:** Completed full repository cleanup for open-source visibility.
- **Liability Protection:** Added explicit disclaimers in `SettingsWindow` and `README.md`.
- **New Feature:** Action Titles Extractor (`BtnActionTitles`) for automated slide headline management.
- **Privacy Mode:** Added "Flight Mode" toggle (`BtnFlightMode`) for quick anonymization.

### Changed
- **Configurable Sidebar:** The Search Bar is now a dynamic widget that can be repositioned or removed via settings.
- **Alignment Refactor:** Consolidated 8 alignment features into a centralized, maintainable execution path.
- **Log Location:** Relocated shortcut debug logs from the Desktop to the system Temp folder for better workspace hygiene.

### Fixed
- **VSTO Stability:** Refactored Slide and Shape Library managers to use asynchronous, STA-safe threading, preventing COM-related crashes.
- **Theming:** Fixed hardcoded colors in `SizeAdjustControl.xaml` to support Light/Dark modes correctly.
- **Local Paths:** Eliminated hardcoded absolute paths in `AutoTaggingService.cs`.

## [2026.4.25.1435] - 2026-04-25

### Changed
- **New Versioning Scheme:** Migrated to yearly versioning (yyyy.mm.dd.hhmm)
- **UI & Layout Overhaul:** Completely restructured the Sidebar into a focused "Consulting Standard" and a comprehensive "All Features" layout
- **Theming:** Full project-wide migration to dynamic theme resources (Light/Dark mode compliant)
- **UX Improvements:** Converted global floating tools to ToggleButtons for better state feedback

### Added
- **Language Support:** Added dynamic Windows system language selection for the Series Generator (Days, Months)
- **Smart Components:** Added EE4P and Think-Cell compatibility for Star Rating, Progress Series, Checkbox, Traffic Light, and Thermometer
- **UI:** New modern WPF-based Checkbox Toolbar replacing the old WinForms implementation

## [1.0.0.4] - 2026-03-12

### Changed
- Restructured Features folder organization
- Added ClickOnce deployment files

## [1.0.0.3] - 2026-03-10

### Added
- AgendaWizard with saved agenda layouts and title formatting
- Dimension lock feature for shape constraints

### Fixed
- Robust slide insertion handling
- ComboBox dropdown menu styling issues

## [1.0.0.2] - 2026-03-08

### Added
- WPF TrafficLight menu and UI styles
- UI\Styles\Controls.xaml and UI\Styles\Themes\
- Thermometer chart feature

### Fixed
- StaticResource dependency bugs on root definitions

## [1.0.0.1] - 2026-03-05

### Changed
- Removed 'E' branding for cleaner interface
- Standardized build Platform configuration

## [1.0.0.0] - 2026-02-28

### Added
- Initial release of oPenEfficiency PowerPoint Add-in

#### Core Features
- Alignment tools (Align to First, Dock, Stretch, Match Size)
- Shape arrangement utilities (Arrange Pro, Arrange Grid, Swap Positions)
- Chart generators (Traffic Light, Harvey Balls, Star Rating, Thermometer)
- Table manipulation (Convert to/from Shapes, Sort, Transpose, Split)
- Text utilities (Split by Paragraphs, Replace Font, Format Numbers)
- Visual enhancements (Pick Color, Theme Color Picker, Remove Gaps)
- Productivity tools (Agenda Wizard, Sticky Notes, Shape Locking, QR Code)
- Asset libraries (Shape, Slide, and Image Library management)

