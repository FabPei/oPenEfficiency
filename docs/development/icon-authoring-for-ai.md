# Author project icons with the minimal Composer

The source checkout includes a headless [icon Composer](../../tools/icon-composer/README.md),
the public subset of approved kit v19, editable JSON documents and a Windows/WPF SVG exporter.
No sibling repository, graphical SVG editor or PowerPoint session is needed.

Use Node.js 22.13+ on the 22.x line or Node.js 24+, npm, Windows PowerShell 5.1 and
.NET Framework 4.8. From the repository root:

```powershell
npm.cmd --prefix tools/icon-composer ci
npm.cmd --silent --prefix tools/icon-composer run editor -- --capabilities
npm.cmd --silent --prefix tools/icon-composer run editor -- --recipe tools/icon-composer/examples/create-slide-add.recipe.json --profile openefficiency --exporter eng/Export-IconComposerSvg.ps1 --output artifacts/icon-composer/new-slide.svg --document artifacts/icon-composer/new-slide.editor.json
```

Use npm.cmd in PowerShell; the npm.ps1 shim can consume the command separator.
On POSIX shells, npm --silent can inspect/edit documents, but the production exporter
requires Windows/WPF. Read the package README for the recipe fields and commands.

Reuse approved unchanged primitives. Keep each recipe with the feature contribution.
An editable SVG/document may retain masks; the production SVG expands strokes and
resolves cutouts for the existing add-in parser. Use semantic colors and inspect
Light/Dark previews at the intended sizes. Explicit manual colors survive theme changes.
The Font and Language primitives are intentionally absent pending their source-rights
clarification; do not silently redraw or reintroduce them.

Run `npm.cmd --prefix tools/icon-composer test` for deterministic CLI and safety
regressions. Follow the package README to set the approved ignored output directories,
generate Light/Dark outputs and run `eng/Test-IconComposerExport.ps1 -ProductionDirectory <recorded-run-directory>`.
That command checks the actual generated SVGs with the production parser and WPF geometry.
The public build workflow runs these authoring checks in addition to the existing
application build/tests and installer generation. An icon-only preview is not real Office
or installation acceptance. See [source notices](../../tools/icon-composer/THIRD_PARTY_NOTICES.md).
