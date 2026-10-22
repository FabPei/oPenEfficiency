# Country flag asset source

- Source: `lipis/flag-icons`
- Release: `v7.3.2`
- Commit: `fe15c16e7463d0c66d6c5730e9d0e832438d98e1`
- Upstream: <https://github.com/lipis/flag-icons>
- License: MIT; see `LICENSE-flag-icons.txt`
- Imported subset: the 249 `country.json` entries with `iso: true`, using `flags/4x3/<alpha-2>.svg`
- Local derivatives: `thumbs/<alpha-2>.png` are 48×36 picker thumbnails rendered from the corresponding SVG.

To regenerate, obtain the pinned upstream version, select its 249 `country.json` entries with `iso: true`, and copy each matching `flags/4x3/<alpha-2>.svg` into this source folder. Render matching SVGs to 48×36 PNG thumbnails in `thumbs/` (for example with CairoSVG). Review the country catalog, generated files and license before accepting another upstream version. The maintainer's import helpers are not part of this public source snapshot.
