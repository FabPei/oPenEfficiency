# Third-party notices

This file describes third-party material distributed with oPenEfficiency. The
project code is licensed under GNU AGPL version 3 (AGPL-3.0-only). The assets
and components listed below retain their separately recorded licenses. Earlier
project-owned MIT asset grants remain available; their license text is bundled
as `Resources/IconLibrary/LICENSES/Project-Assets-MIT.txt`. AGPL does not
replace those independent terms or previously valid grants.

## Icon assets

### Searchable icon library

The packaged icon-library archive and manifest contain the following upstream
collections. Exact counts, copyright notices, source links, and the known
snapshot limitations are recorded in
[`oPenEfficiency/Resources/IconLibrary/THIRD_PARTY_NOTICES.md`](oPenEfficiency/Resources/IconLibrary/THIRD_PARTY_NOTICES.md).
The corresponding license texts are distributed in
`oPenEfficiency/Resources/IconLibrary/LICENSES/`.

| Collection | License |
|---|---|
| Phosphor Icons | MIT |
| Tabler Icons | MIT |
| Lucide Icons | ISC; Feather-derived icons remain MIT |
| Iconoir | MIT |
| Heroicons | MIT |

### Standalone application icons

`oPenEfficiency/UI/Assets/Icons/` contains 201 SVG files, of which 197 are
included in the application payload and four are retained as unused development
assets. Their provenance is classified as follows:

- 86 files carry Lucide class metadata; 83 of these are shipped. They are
  modified or recolored Lucide assets under the ISC license, except for the
  Feather-derived icon names identified in
  `Lucide-ISC-and-Feather-MIT.txt`. Three additional Phase 7 selections without
  retained class metadata add
  `BtnChess.svg`, `BtnLocalSlideRebuild.svg`, and
  `BtnMatchAngles.svg` from the bundled Lucide catalog. The 2026-10-20
  `BtnThemeColor.svg` painter-palette replacement (owner direction: a palette
  look from Lucide/Phosphor) is a renamed Lucide `palette` selection without
  retained class metadata, replacing the former Phosphor eyedropper geometry.
- The owner-selected 2026-10-20 `BtnExploreFeatures.svg` replacement composes only
  project-licensed kit primitives (Collection main object with the Inspect magnifier
  as `ActionBlue` action modifier, sibling `svg_editor` Composer export per
  `docs/product/icon-proposals/2026-10-20-feature-explorer/`). It replaces the former
  Lucide question-mark geometry, reducing the metadata-bearing Lucide inventory by
  one; no third-party geometry remains in this file.
- The owner-directed 2026-10-20 `BtnTableSort.svg` replacement composes the
  project Table primitive (`docs/product/icon-building-blocks/objects/table.svg`)
  with the former Lucide `arrow-down-az` action geometry retained as a
  bottom-right ActionBlue modifier, authored through the sibling `svg_editor`
  Composer and exported via `eng/Export-IconComposerSvg.ps1`. Replacing the plain
  Lucide `arrow-down-az` selection reduces the metadata-bearing Lucide inventory
  by one; the composed derivative remains explicitly attributed under ISC and its
  recipe is recorded in
  `docs/product/icon-proposals/2026-10-20-table-sort/`.
- The owner-selected 2026-10-19 replacement for `BtnAnonymize.svg` combines the
  project text-box primitive with Lucide Hat + Glasses geometry under the ISC license.
  The four Excel-family replacements reuse owner-selected project primitives.
  `BtnFormatNumbers.svg` now uses the owner-selected 2026-10-22 Noto Sans SemiBold
  `123` outlines, replacing the prior Segoe UI source. Noto Sans is licensed under
  SIL OFL 1.1; the selected graphic was generated from its original contours.
  Current number-format source hashes are retained in
  `docs/product/icon-proposals/2026-10-22-ofl-symbol-replacements/production/integration.json`;
  the earlier six-icon record retains historical provenance. Removing the Euro badge reduces the
  metadata-bearing Lucide inventory by one; the unmarked Hat + Glasses derivative
  remains explicitly attributed here and in the compliance inventory.
- The owner-selected `BtnDeleteText.svg` 1B replacement reuses the approved project
  text-box and delete primitives, replacing its previous Lucide-derived geometry.
  Exact source hashes are recorded in
  `docs/product/icon-proposals/2026-10-19-delete-text/selection.json`.
- The seven owner-selected family replacements of2026-10-19 reuse approved project
  primitives. The Omega/Pi special-character source was subsequently replaced on
  2026-10-22 by owner-selected Noto Sans SemiBold outlines under SIL OFL 1.1.
  `BtnSlideNotes.svg` includes original Lucide Hat/Glasses
  within the approved private-note master; `BtnSlideSize.svg` reuses the original
  Lucide Weight paths as a modifier. Both remain explicitly attributed under ISC
  despite the absence of class metadata. `BtnConnectedTextLabel.svg` now uses
  project Text Box/Link primitives, replacing its former Iconoir source. Exact
  provenance is in docs/product/icon-proposals/2026-10-19-family-variants/selection.json,
  provenance.json and the referenced symbol recipes.
- 23 shipped files match Phosphor icon geometry and are renamed or recolored
  derivatives under the MIT license: `arrow-fat-lines-right.svg`,
  `arrow-fat-lines-up.svg`, `BtnArrangeGrid.svg`,
  `BtnFlipHorizontal.svg`, `BtnFlipVertical.svg`,
  `BtnFormatItalic.svg`, `BtnFormatUnderline.svg`, `BtnInsertColumnLeft.svg`,
  `BtnInsertColumnRight.svg`, `BtnInsertRowBottom.svg`,
  `BtnInsertRowTop.svg`, `BtnMagicResizer.svg`,
  `BtnMapWizard.svg`, `BtnOptimizeFreeForm.svg`, `BtnReplaceFont.svg`,
  `BtnReplaceShape.svg`, `BtnTableBranding.svg`,
  `BtnTableColumnInsertion.svg`, `BtnTableRowInsertion.svg`,
  `BtnTrafficLight.svg`,
  `BtnTranslate.svg`, `BtnTransparency.svg`, and `colorpalette.svg`.
- `BtnTextDirection.svg` is a renamed Iconoir
  derivative under the MIT license. `BtnDocumentAutomation.svg` is a renamed
  Tabler Icons derivative under the MIT license.
- Seven Phase 7 icons are project-specific drawings created for this repository:
  `Btn2048.svg`, `BtnExcelLinkManager.svg`, `BtnNewSlide.svg`,
  `BtnPositionPainter.svg`, `BtnSeriesGenerator.svg`,
  `BtnSplitByParagraphs.svg`, and
  `BtnTextFormatPainter.svg`.
- The ten owner-selected Ribbon redraws integrated on 2026-10-10 are original
  project SVG drawings: `BtnBringForward`, `BtnSendBackward`, `BtnHideMasterObjects`,
  `BtnConvertTableToShapes`, `BtnTableColumnWidth`, `BtnTableRowHeight`,
  `BtnTableDimensions`, `BtnLearnMargin`, `BtnAddSticker`, and `BtnSlidePaste`.
  Their exact source files and hashes are recorded in
  `docs/product/icon-proposals/2026-10-10-last-three-variants/selection.json`.
- A further three original project drawings were owner-selected on 2026-10-10:
  `BtnArrangeCircle`, `BtnArrangeInShape`, and `BtnChartOverlay`. Five SVGs supplied
  by the owner in `Temp folder New Icons` replace `BtnSnapToGrid`, `BtnAssetManager`,
  `BtnGanttChart`, `BtnAlignToTableCell`, and `BtnTableHeatmap`. Their original files, source hashes and
  renderer normalization are recorded in
  `docs/product/icon-proposals/2026-10-10-selected-eight/selection.json`.
- Eight further owner-selected original project drawings replace the six
  Increase/Decrease Spacing and Remove Gaps commands, `BtnShapeLock` and
  `BtnTagInspector`. Exact sources and hashes are recorded in
  `docs/product/icon-proposals/2026-10-10-fourth-ten/selection.json`.
- Three further original project drawings replace `BtnProgressSeries`,
  `BtnIconBrowser` and `BtnTemplateManager`, exactly matching the owner-selected
  fifth-batch SVGs. Their source hashes are recorded in
  `docs/product/icon-proposals/2026-10-10-fifth-ten/selection.json`.
- The remaining application icons are project-specific compositions or drawings
  first recorded in this repository and are distributed under their recorded
  MIT license (see `Resources/IconLibrary/LICENSES/Project-Assets-MIT.txt`). No external source is identified in repository history. This
  includes `BtnIllustrationLibrary.svg`, drawn for the Illustration Library
  feature (FEAT-011), and `BtnFeatureAcceptanceWindow.svg`, drawn for the
  developer/QA acceptance status window (QA-S6).

The searchable icon library's `custom` directory contains another 65 SVG files
generated with an AI image-generation system at the direction of the project
maintainer. The specific model, prompts, and per-file generation dates were not
retained. No externally sourced icon files are known to have been imported into
this collection. To the extent that copyrightable human selection, arrangement,
or modification is present and controlled by the project maintainer, those
contributions retain their recorded MIT license; the text is bundled as
`Resources/IconLibrary/LICENSES/Project-Assets-MIT.txt`. The project does
not claim exclusive copyright in purely AI-generated elements. An externally
sourced replacement or addition must record its author, source, revision or
creation date, and license before it is accepted.

### Noto Sans outline artwork

`UI/Assets/Icons/BtnFormatNumbers.svg` and `BtnSpecialChars.svg` use the owner-selected
number-C and symbol-C outline artwork from Noto Sans SemiBold, Copyright 2018 The
Noto Project Authors, under SIL Open Font License 1.1. The original font contours
were fitted mechanically to the icon canvas. The font software itself is not
bundled with the application or installed by it. The source font, unmodified OFL
text, exact revision and authoring recipes are retained in the private development
record `docs/product/icon-proposals/2026-10-22-ofl-symbol-replacements/`.

Noto source: https://github.com/notofonts/noto-fonts/tree/ffebf8c1ee449e544955a7e813c54f9b73848eac

OFL FAQ 1.1–1.1.2 permits graphics made from font outlines and does not require
those resulting graphics to be licensed under OFL:
https://openfontlicense.org/ofl-faq/

### Country flag assets

`oPenEfficiency/Resources/Flags/` contains 249 ISO 3166-1 country flag SVGs and
local picker thumbnails imported from `lipis/flag-icons` v7.3.2 (commit
`fe15c16e7463d0c66d6c5730e9d0e832438d98e1`). The collection is distributed
under the MIT license; the upstream license is bundled as
`oPenEfficiency/Resources/Flags/LICENSE-flag-icons.txt`. PNG files are local
renderings of the corresponding SVG artwork and retain the same provenance.

## Geographic data

### Natural Earth

The Admin 0, Admin 1, world, and continent assets whose filenames contain
`naturalearth_5_1_1` are derived from Natural Earth 5.1.1. Natural Earth has
made these vector data available in the public domain:

https://www.naturalearthdata.com/about/terms-of-use/

oPenEfficiency projects, simplifies, normalizes, and converts the source
geometry for editable PowerPoint output. Natural Earth does not endorse this
project.

### geoBoundaries gbOpen

The global ADM2 and ADM3 assets whose filenames contain
`geoboundaries_current` are modified and simplified derivatives of
geoBoundaries gbOpen. geoBoundaries distributes gbOpen under Creative Commons
Attribution 4.0 International (CC BY 4.0).

- Creator/data provider: geoBoundaries, William & Mary geoLab and community
  contributors
- Source: https://www.geoboundaries.org/
- License: https://creativecommons.org/licenses/by/4.0/
- Suggested citation: Runfola, D. et al. (2020), “geoBoundaries: A global
  database of political administrative boundaries,” PLOS ONE 15(4): e0231866,
  https://doi.org/10.1371/journal.pone.0231866
- Changes made by oPenEfficiency: projection, precision reduction,
  simplification, normalization, attribute removal, parent-region assignment,
  compression, and conversion to the application's map schema.

### German ADM2 and ADM3 data

`de_adm2_government_districts_geoboundaries_2021.json.gz` and
`de_adm3_districts_geoboundaries_2021.json.gz` are modified derivatives of the
2021 Germany datasets distributed through geoBoundaries. The source metadata
identifies the Federal Agency for Cartography and Geodesy (BKG) and the “Data
licence Germany – attribution – Version 2.0” (`dl-de/by-2-0`).

- Data source: Federal Agency for Cartography and Geodesy (BKG), 2021
- Distribution/source metadata: geoBoundaries
- License information: https://www.govdata.de/dl-de/by-2-0
- Changes made by oPenEfficiency: projection, simplification, normalization,
  attribute removal, naming/alias enrichment, compression, and conversion to
  the application's map schema.

## Build-time and runtime dependencies

The test projects use Microsoft.NET.Test.Sdk and MSTest under MIT. Roslyn's
Microsoft.CodeAnalysis.NetAnalyzers package is also used as a private build-time
dependency under MIT. These packages are not part of the application payload.

The VSTO projects use Office interop assemblies as compile-time references.
The NuGet packages used to provide those reference assemblies do not declare an
open-source license and are not represented as Microsoft-supported packages.
Consequently, their runtime assets are explicitly excluded from the
oPenEfficiency distribution and only the required interop type information is
embedded in the add-in assemblies (NoPIA). At runtime the add-ins rely on the
Office/VSTO components installed and licensed on the user's machine.

`Microsoft.Office.Tools.Common.v4.0.Utilities.dll` is distributed unmodified
with each add-in because Microsoft identifies this VSTO utility assembly as a
Visual Studio redistributable and its VSTO deployment guidance requires utility
assemblies referenced by a solution to be deployed with that solution. Its use
and redistribution remain subject to the applicable Microsoft Visual Studio
license terms:

https://learn.microsoft.com/visualstudio/releases/2026/redistribution

Microsoft, Microsoft Office, PowerPoint, Excel, Visual Studio, and VSTO are
trademarks of the Microsoft group of companies. Their mention does not imply
endorsement.

## Source provenance references

Paths under `docs/product/icon-proposals/` and `docs/product/icon-building-blocks/` in historical provenance descriptions refer to private authoring records, not public download locations. The distributed SVG geometry, upstream attributions and license texts are included in this public source. Those records do not change the independent third-party terms.
