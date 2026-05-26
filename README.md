# oPenEfficiency

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Build](https://github.com/FabPei/oPenEfficiency/actions/workflows/build.yml/badge.svg)](https://github.com/FabPei/oPenEfficiency/actions/workflows/build.yml)
[![Release](https://img.shields.io/github/v/release/FabPei/oPenEfficiency)](https://github.com/FabPei/oPenEfficiency/releases)

> **Important:** oPenEfficiency is a personal, experimental collection of productivity tools for Microsoft PowerPoint. It is shared as-is with the open-source community as a hobby project.

oPenEfficiency is a **Microsoft PowerPoint VSTO Add-in** built to bridge the gap between standard PowerPoint features and the high-speed requirements of professional consulting and design workflows.

As a personal collection of tools, it focuses on eliminating the repetitive, manual "pixel-pushing" that consumes valuable time during presentation building. Whether it's ensuring 20 shapes are perfectly aligned to a specific anchor, transforming messy text into structured tables, or maintaining a consistent narrative flow with automated agendas, oPenEfficiency provides a "Swiss Army Knife" of automation directly within your PowerPoint ribbon.

The toolkit is designed with a "Consultant-First" mindset, prioritizing features like:
- **Zero-Effort Alignment:** Tools that understand context, like "Align to First" and "Move Until Collision."
- **Smart Data Visuals:** Lightweight, editable vector components like Harvey Balls and Star Ratings that feel native to PowerPoint.
- **Table Mastery:** Solving the long-standing frustration of PowerPoint tables with specialized formatting and structural manipulation tools.
- **Workflow Automation:** One-click solutions for agenda generation, style auditing, and branding anonymization.

Built on the .NET Framework and deep PowerPoint COM integration, oPenEfficiency offers a responsive, native-feeling experience while remaining highly customizable through its integrated Settings and Sidebar system.

![PowerPoint Add-in](https://img.shields.io/badge/PowerPoint-2016%2B-BD41A0?logo=microsoft-powerpoint)
![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.8-512BD4?logo=dotnet)

---

## ⚠ Disclaimer

**This is a hobby project and not a commercial product.** It is provided **WITHOUT ANY WARRANTY**, express or implied. The author assumes no liability for any data loss, performance issues, or other problems caused by the use of this software. Use it at your own risk.

---

## Features

### 📐 Alignment & Layout
- **Align to First** - Align selected shapes to the first selected shape (left, right, top, bottom, center).
- **Stack & Dock** - Instantly stack shapes adjacent to each other or dock them to slide/shape edges.
- **Stretch & Match** - Stretch shape edges to match an anchor or standardize Width/Height across a selection.
- **Swap Positions** - Swap locations of two shapes relative to customizable anchor points (Center, Top-Left, etc.).
- **Multi-Swap** - Swap multiple shapes based on selection order and direction (Top-Down, Bottom-Up, etc.).
- **Arrange Tools** - Arrange shapes in perfect grids, circles, triangles, or complex "Pro" patterns (spirals, concentrics).
- **Precision Spacing** - Increase/decrease horizontal/vertical gaps by fixed increments or remove gaps entirely.
- **Diagonal Align** - Position shapes along a diagonal axis.
- **Position Painter** - Copy and apply the exact coordinates and size from one shape to another.
- **Rectify Rotation** - Reset rotation to 0 or snap to the nearest 90-degree increment.
- **Table Cell Alignment** - Precision alignment of shapes within PowerPoint table cells.

### 🎨 Visual Elements & Formatting
- **Infographics** - Generate and manage **Harvey Balls**, **Star Ratings**, **Thermometers**, **Traffic Lights**, and **Checkboxes**.
- **Glass Hide** - Create semi-transparent overlays to obscure content or add frosted glass effects.
- **Progress Series** - Create automated chevron/arrow sequences for process steps.
- **Smart Components** - Synchronized **Smart Corners** for multiple shapes and **Align Shape Proportions** for consistency.
- **Color Tools** - High-speed **Theme Color Palette**, global **Color Picker (Eyedropper)**, and **Color Eraser** (Transparency).
- **Object Connector** - Create dynamic, locked connectors between shapes that follow movement.
- **Advanced Formatting** - **Optimize Free Form** (smooths hand-drawn paths) and **Bulk Format Replacer** (batch update fonts/colors).
- **Stickers & Tags** - Insert professional status stickers (WIP, Confidential) or custom **Numbered Bullets**.
- **Sync Objects** - Keep properties (size, color, text) synchronized across multiple shapes.

### 📊 Table Productivity
- **Format Painter** - A "Format Painter" designed specifically for table cells (Up, Down, Left, Right).
- **Structural Tools** - **Transpose** (flip rows/columns), **Split Table**, and **Sum Table Cells** (adds automatic totals).
- **Table Intelligence** - **Table Heatmap** for data visualization and **Table Sort** for organizing rows.
- **Conversion** - Convert any selection of shapes into a native table or explode a table back into editable vector shapes.
- **Dimension Control** - Precise control over column widths and row heights with copy/paste functionality.
- **Table Branding** - Apply consistent company styling to existing tables with one click.

### 📝 Text Utilities
- **Structure** - **Split by Paragraphs** (distribute text into multiple shapes) or **Merge Text**.
- **Translation** - Integrated translation using **DeepL**, **LibreTranslate**, or **LLM** services directly on slides.
- **Numerical** - **Number Formatter** for standardized decimal/thousands styling.
- **Transformation** - **Swap Text** (exchange content between shapes), **Change Case**, and **Apply Text to Selection** (batch edit).
- **Formatting** - **Align Text**, **Fit Form to Text**, **Delete Text**, and quick access to **Spell Check Languages**.

### 🚀 Storylining & Review
- **Agenda Wizard** - Full-featured system for generating and maintaining presentation agendas and layouts.
- **Action Titles** - Extract and review all slide headlines in a single view to check narrative flow.
- **Excel Link Manager** - Manage and bulk-update links to Excel charts and data ranges.
- **Review Tools** - **Style Checker** (scans for inconsistencies), **Slide Guidelines** (layout grids), and **Cleaner** (strips hidden data).
- **Sticky Notes** - Manager for internal comments and "off-slide" annotations.
- **Flight Mode** - Privacy toggle to temporarily obscure logos and branding during public presentations.

### 🛠 Utilities & Wizards
- **Asset Library** - Searchable repository for your custom **Shapes**, **Slides**, and **Images**.
- **Snap Tools** - Advanced **Snap to Grid** and **Snap to Objects** for pixel-perfect placement.
- **Selection & Locking** - **Select Same Type**, **Shape Locking**, and **Dimension Lock** to protect your layout.
- **Coordinate Tools** - **Copy/Paste Coordinates** for precise positioning across different slides.
- **Visibility Manager** - **Hide Selected**, **Show Hidden**, and **Hide Master Objects** to focus on what matters.
- **Automation** - **Export Wizard**, **Series Generator**, **Map Wizard**, and **Document Automation**.
- **Technical** - **Tag Inspector** for viewing hidden shape metadata and **Motion Path Creator**.

---

## Installation

### Prerequisites
- Windows 10/11
- Microsoft PowerPoint 2016 or later
- .NET Framework 4.8
- [Visual Studio 2022 Tools for Office Runtime](https://aka.ms/vstor_100_150)

### Manual Installation (Development)

1. Clone or download this repository.
2. Open `oPenEfficiency.slnx` in Visual Studio 2022.
3. Ensure the "Office/SharePoint development" workload is installed.
4. Build the solution in **Release** configuration.
5. The add-in will be registered on your machine. Restart PowerPoint to see the new ribbon.

---

## Usage

Once installed, oPenEfficiency adds a custom "oPen Efficiency" tab to the PowerPoint ribbon.

1. **Select shapes** or elements on your slide.
2. **Click a feature** in the ribbon or sidebar.
3. Configure settings in the **Settings Window** (Gear icon) to customize your experience.

---

## Sidebar & UI Customization

oPenEfficiency is designed to be highly personalizable. The Sidebar serves as your primary productivity hub and can be tailored to match your specific workflow.

### 🛠 Sidebar Layout
- **Custom Sections** - Create, rename, or remove sections to group your most-used tools.
- **Feature Management** - Add or remove any button from the sidebar. You can also reorder features within sections to optimize your "click-path."
- **Layout Presets** - Choose from built-in presets like **Consulting Standard**, **All Features**, or **Minimalist**.
- **Visibility Toggle** - Instantly hide entire sections you don't need for specific projects.

### 🎨 Visual Styling
- **Themes** - Switch between **Dark Mode** and **Light Mode** to match your PowerPoint environment.
- **Custom Colors** - Set specific hex colors for the Sidebar background, Section headers, and Button backgrounds.
- **Typography** - Adjust the **Font Family**, **Section Font Size**, and **Global Icon Size** for better legibility on high-DPI displays.
- **Transparency** - Enable transparent backgrounds to let the Sidebar blend seamlessly into the PowerPoint workspace.

### ⚙ Behavioral Settings
- **Anchor Logic** - Toggle between "First Selected" and "Last Selected" as the default anchor for all alignment tools.
- **Incremental Control** - Define custom point-increments for the "Precision Spacing" and "Magic Resizer" tools.
- **Sticker Defaults** - Customize the list of standard status stickers (e.g., "Draft," "Client Review") and their default colors/fonts.
- **Sticky Note Styles** - Configure default fonts and colors for annotations to ensure they stand out or remain subtle.

### 📤 Configuration & Sharing
- **Full Export/Import** - Your entire configuration (shortcuts, layout, and colors) can be exported to a single JSON file.
- **Layout Sharing** - Export *only* your sidebar layout to share an optimized workflow with your team without overwriting their personal color settings.
- **Emergency Reset** - One-click "Reset to Defaults" to restore the factory consulting-standard layout.

---

## Keyboard Shortcuts
Default shortcuts (Customizable in Settings):
- `Ctrl+Alt+S` - Toggle Sidebar
- `Ctrl+Alt+M` - Match Size
- `Ctrl+Alt+A` - Align to First

---

## Development

### Project Structure
```
oPenEfficiency/
├── Features/           # Core feature logic (Alignment, Tables, Visuals, etc.)
├── Models/             # Data structures and configuration models
├── Services/           # Business logic and PowerPoint interaction layer
├── UI/                 # WPF dialogs, sidebar panels, and custom controls
└── Utils/              # Helper classes and Win32 hooks
```

---

## License

This project is licensed under the GNU GPL 3.0 License - see the [LICENSE](LICENSE) file for details.

---

## Acknowledgments

This project is built using Microsoft's VSTO (Visual Studio Tools for Office) framework. 

**Icon Credits:**
- Icons are sourced from the excellent open-source libraries: [Phosphor Icons](https://phosphoricons.com/), [Lucide](https://lucide.dev/), and [Hugeicons](https://hugeicons.com/).
- Design elements are inspired by standard consulting productivity tools.
