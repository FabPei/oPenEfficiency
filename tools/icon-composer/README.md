# openEfficiency Icon Composer for coding AIs

This source-repository tool provides the approved Composer engine and the
74-source public subset of project kit v19, a JSON recipe format, and editable document plus
production SVG output. It runs from a fresh openEfficiency checkout and does
not require the sibling `svg_editor` checkout or its React/Vite application.

## Prerequisites

- Windows 10 or newer, Node.js 22.13 or newer within the 22.x line, or Node.js 24 or newer, and npm.
- Windows PowerShell 5.1 and .NET Framework 4.8 for the production profile.
- No PowerPoint installation is needed for command-line authoring/export.

From the repository root:

```powershell
npm.cmd --prefix tools/icon-composer ci
npm.cmd --silent --prefix tools/icon-composer run editor -- --capabilities
npm.cmd --silent --prefix tools/icon-composer run editor -- --recipe tools/icon-composer/examples/create-slide-add.recipe.json --profile openefficiency --exporter eng/Export-IconComposerSvg.ps1 --output artifacts/icon-composer/new-slide.svg --document artifacts/icon-composer/new-slide.editor.json
```

An editable native document can be reopened and modified with `--input` and a
`--commands` JSON file. `--commands -` reads a command array from stdin. Recipe
JSON supports `id`, `name`, approved kit IDs in `main` and `action`, and partial
Composer `params`. Keep the recipe in source control; the native document keeps
the editable layers and Composer metadata. The production SVG is the WPF-ready
output. Output paths are explicit and may be overwritten.

When `baseColor` is omitted or empty, Composer resolves the main foreground from
`tokenTheme` for recipe and command `compose` operations. Editable metadata keeps
that theme-derived default empty so `recompose` can follow later theme changes.
A supplied `baseColor` remains manual across theme changes, including when it
equals a theme's default color. Existing documents with a concrete `baseColor`
are conservatively treated as manual because their history cannot be inferred;
set `baseColor` to an empty string to resume theme-derived color.

## Machine interface

`--capabilities` returns JSON with versioned operations, Composer action
positions, the kit version/count and recipe fields. Success is JSON on stdout
with `ok: true`, command results, document snapshot, SVG and written file paths.
Errors are JSON on stderr with `ok: false` and a nonzero exit status. A failed
command batch rolls back; failed WPF export writes no requested output files.

For direct document editing:

```powershell
npm.cmd --silent --prefix tools/icon-composer run editor -- --input artifacts/icon-composer/new-slide.editor.json --commands edits.json --profile openefficiency --exporter eng/Export-IconComposerSvg.ps1 --output artifacts/icon-composer/edited.svg --document artifacts/icon-composer/edited.editor.json
```

Run the bundled deterministic CLI regressions with
`npm.cmd --prefix tools/icon-composer test`. To also export Light/Dark bundle outputs and
run the real production parser/WPF geometry gate against those exact outputs:

```powershell
$env:OE_ICON_EXPORTER = (Resolve-Path eng/Export-IconComposerSvg.ps1).Path
$env:OE_ICON_PRODUCTION_DIR = (Join-Path (Get-Location) 'artifacts/icon-composer-bundle/production')
$env:OE_ICON_ARTIFACTS_DIR = (Join-Path (Get-Location) 'artifacts/icon-composer-bundle/previews')
$env:OE_ICON_EVIDENCE_PATH = (Join-Path (Get-Location) 'artifacts/icon-composer-bundle/light-dark-size.json')
npm.cmd --prefix tools/icon-composer test
$bundleRun = Get-Content artifacts/icon-composer-bundle/light-dark-size.json -Raw | ConvertFrom-Json
eng/Test-IconComposerExport.ps1 -ProductionDirectory $bundleRun.outputDirectories.production
Remove-Item Env:OE_ICON_EXPORTER,Env:OE_ICON_PRODUCTION_DIR,Env:OE_ICON_ARTIFACTS_DIR,Env:OE_ICON_EVIDENCE_PATH
```

The exporter gate rebuilds the net48 test host and runs its deterministic unit checks,
then parses the new bundle SVG files through the add-in's production `SvgParser` and
`FeatureIconPathFactory`. Each test execution creates a fresh `run-*` child under the
approved ignored artifact directories. Previous outputs and unrelated marker files are
left untouched; the exact production directory is recorded in the execution JSON.

## Provenance and rights

The maintainer confirmed ownership of the extracted editor code and its delivery
under the existing openEfficiency AGPL v3 license. The root project license is
unchanged. Separately licensed icons, font-derived graphics and dependencies retain
their own terms; see [third-party notices](THIRD_PARTY_NOTICES.md).

`source-manifest.json` records the bounded payload and hashes. Private source
history, dirty-working-tree inventories and design-review boards are not included.
The approved kit retains original SVG geometry and hashes. The Font and Language
building blocks are excluded from this public subset because reuse rights for
their Times New Roman glyph outlines have not been established. Do not recreate
those missing outlines by hand. The Noto Sans number graphic has its pinned
OFL provenance included. No font binary or full graphical editor is shipped.

Third-party runtime dependencies are pinned in `package-lock.json`:

- `jsdom` and `tsx`: MIT; upstream license texts are included in the installed
  npm package metadata.
- `@resvg/resvg-js`: MPL-2.0; upstream license text is included in the
  installed npm package metadata. The package is used only to validate/rasterize
  requested previews; the production SVG is generated by the existing project
  PowerShell/.NET exporter.

The approved kit includes reused Lucide, Microsoft Fluent, Material Design
Icons, and Noto Sans outlines where recorded by kit provenance. Corresponding
Lucide, MIT, Pictogrammers, Apache-2.0 and SIL OFL 1.1 notices and source
provenance records are retained under `kit/notices/` and `kit/provenance/`.

No dependency source is copied into this repository. Preserve these notices
when redistributing the tool or adapted sources.
