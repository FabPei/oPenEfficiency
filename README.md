# openEfficiency

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](LICENSE)
[![Release](https://img.shields.io/github/v/release/FabPei/oPenEfficiency)](https://github.com/FabPei/oPenEfficiency/releases)

## Your consulting toolkit, inside PowerPoint

**Build sharper slides. Spend less time formatting.** openEfficiency is an open-source PowerPoint add-in for consultants, analysts, strategy teams and anyone who turns complex ideas into professional presentations.

From the first working slide to the final client deck, it brings a broad set of presentation tools into one configurable workspace: precise alignment and spacing, table and text tools, Harvey balls and ratings, agendas, reusable assets and linked Excel chart updates. These are the small, frequently used tools that make a difference when you are structuring a recommendation, polishing a pitch or working through last-minute revisions.

**Less pixel-pushing. More time for the story.** Arrange a crowded slide, make a table consistent, clean up text and keep your deck moving—all from PowerPoint's ribbon and a customizable sidebar. Choose the tools that fit your workflow and keep them close at hand.

Built for Windows desktop PowerPoint with VSTO and .NET Framework 4.8. Maintained as a personal open-source project and free to use under the AGPL.

[Download a release](https://github.com/FabPei/oPenEfficiency/releases) · [Installation](https://github.com/FabPei/oPenEfficiency/wiki/Installation) · [User guide](https://github.com/FabPei/oPenEfficiency/wiki/Usage) · [Feature reference](https://github.com/FabPei/oPenEfficiency/wiki/Features) · [Report a bug](https://github.com/FabPei/oPenEfficiency/issues/new/choose)

## Current public status

The current release is [openEfficiency V0.9 - 26.10.23](https://github.com/FabPei/oPenEfficiency/releases/tag/v26.10.23). It provides x86/x64 MSI installers and matching SHA-256 files. The public source branch and a downloaded release can differ; use the source tag and instructions for the version you intend to build or install.

This documentation update does not publish a new product version. Features described in the source reference may differ from the build you have installed. Check the [release notes](https://github.com/FabPei/oPenEfficiency/releases) for your version.

## What it helps with

| Area | Examples in the public feature reference |
| --- | --- |
| Alignment and layout | Align, distribute, stack, match size and arrange shapes |
| Tables | Adjust row/column dimensions, transpose, sort and format cells |
| Text | Format numbers, split or merge text, change case and fit shapes to text |
| Visual elements | Harvey balls, ratings, progress elements and color tools |
| Presentation workflow | Agendas, title review, linked Excel charts and reusable assets |
| Workspace | Configurable sidebar, themes and keyboard shortcuts |

See the [detailed feature reference](https://github.com/FabPei/oPenEfficiency/wiki/Features). Command availability, names and options depend on the installed version. This is a PowerPoint add-in; linking to Excel content does not mean a separate Excel add-in is included in these public downloads.

## Current source workspace

The source snapshot adds a PowerPoint **Home** tab launcher and a customizable sidebar as the primary PowerPoint command surface. PowerPoint intentionally uses the Home-tab Sidebar launcher. Choose **Consulting Standard**, **All Features**, **Minimal** or **Labs** in Settings. The source configuration ships with no default feature shortcuts; assign your own in Settings.

The optional Excel companion uses a dedicated **openEfficiency** Excel ribbon tab. Its source is included here; select it as an optional component in the current MSI; older June 2026 downloads use a different package.

The complete [source feature reference](docs/Features.md) covers the current commands and widgets. Examples include **Insert Linked Text Label**, which connects a PowerPoint label to a supported source, and chart overlays for a selected native chart or one group containing exactly one native chart. Outputs use native PowerPoint shapes, tables and text where supported; color erasing deliberately replaces a shape with a raster image.

**Preview:** Treat any feature marked Preview in an installed build as provisional. **Experimental:** Labs tools are experimental. Source documentation and automated checks do not establish Office, visual or release acceptance. OneNote host projects are excluded; existing shared compatibility code remains. Additional Gantt work remains deferred.

### Build this source

Use Windows, Visual Studio with Office/VSTO tools and the .NET Framework 4.8 SDK, a current .NET SDK and PowerShell 7 (`pwsh`). Run `pwsh -NoProfile -File ./eng/Build.ps1 -Configuration Release -Rebuild -Installer` for the source build, five test projects, available host-free acceptance checks and local x86/x64 installer generation. It suppresses live Office add-in registration, restores temporary VSTO development metadata after the build, and does not install MSIs or publish a release. Local VSTO manifest signing uses a temporary self-signed key; it is not a trusted publisher certificate. MSI packages remain unsigned. The maintainer's private release/host evidence is separate.

## Installation

You need Windows desktop PowerPoint, .NET Framework 4.8 and the Microsoft VSTO runtime. PowerPoint 2016 and later is the project's compatibility target; this is not a claim that every Office/Windows combination has been verified. PowerPoint for the web and PowerPoint for macOS do not load Windows VSTO add-ins.

1. Open [Releases](https://github.com/FabPei/oPenEfficiency/releases) and choose a published version.
2. Read the release notes and download the MSI matching your **Office architecture** (x86 for 32-bit Office, x64 for 64-bit Office), plus its `.msi.sha256` file. Compare the SHA-256 hash before installation.
3. Save your work, close PowerPoint and Excel, then run the selected MSI and follow its setup instructions. Older ClickOnce releases retain their own version-specific instructions.
4. Reopen PowerPoint and look for the add-in's ribbon tab or sidebar.

All openEfficiency MSI releases are intentionally unsigned; Authenticode signing is not planned. Windows may show **Unknown publisher** or SmartScreen warnings, and managed devices may block installation. Published SHA-256 files verify bytes, not publisher identity. Technically required self-signed VSTO manifest signatures do not sign the MSI. Follow your organization's policy; do not disable Office security settings to force installation. See [Installation and troubleshooting](https://github.com/FabPei/oPenEfficiency/wiki/Installation).

Developers can build the public source with a Visual Studio version that supports the project format and has Office/VSTO development tools and the .NET Framework 4.8 SDK. A successful compile is not an installation procedure. See [Development](https://github.com/FabPei/oPenEfficiency/wiki/Development).

## Everyday use

Select the shapes, text or table cells you want to change, then choose a command from the ribbon or sidebar. Selection order can determine the alignment anchor. Check the command's tooltip and available context-menu options before applying it to a large selection.

Use Settings to configure the sidebar, visual preferences and shortcuts. The shortcuts currently configured in your installation are the useful reference; this page does not assume that every installation has the same defaults. Try unfamiliar commands on a copy of a presentation and review the result.

Most editing commands work inside PowerPoint. Optional translation integrations can send selected text to the configured service and can require an endpoint or API key. Check the provider's terms before sending confidential content. See [Configuration](https://github.com/FabPei/oPenEfficiency/wiki/Configuration).

## Help and contributions

For bugs, include the add-in version, Windows and PowerPoint versions, Office bitness, selection details and steps to reproduce. Remove customer data, file paths and credentials from logs or screenshots before posting them publicly.

[Open an issue](https://github.com/FabPei/oPenEfficiency/issues/new/choose) · [Contribution guide](CONTRIBUTING.md) · [Code of conduct](CODE_OF_CONDUCT.md)

Maintenance depends on the maintainer's availability; there is no guaranteed support response or service-level agreement. The software is provided without warranty to the extent allowed by its license and applicable law.

## License and resale

openEfficiency is licensed under **AGPL v3**. Private use, business use, modification, free redistribution and selling copies are permitted under its conditions.

A distributor must retain required notices, keep covered modifications under AGPL and provide the matching source as required by the license. Buyers retain their AGPL rights, including the ability to share their copy free of charge. Redistribution does not require paying the maintainer a royalty or contributing changes back to this repository.

Read the [practical licensing guide](docs/licensing.md) and the full [LICENSE](LICENSE). Separately licensed components and artwork retain their own terms.
