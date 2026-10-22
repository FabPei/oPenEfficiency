# Legacy personal library fixture (AM-S0)

Deterministic fixture mirroring the on-disk layout of the legacy personal asset library
root `%APPDATA%\oPenEfficiency\favorites\` as frozen by
`docs/testing/feat-008-current-asset-manager-compatibility.md`.

## Purpose

AM-S1/AM-S2 Core tests (manifest, migration, indexing, search) must run without Office.
This fixture provides the deterministic subset of the legacy format those tests can rely
on. PowerPoint binaries cannot be checked in deterministically and are intentionally
absent (see `slides/README.txt`); tests that need real `.pptx` containers belong to the
owner-executed real-host evidence path, not to Core fixtures.

## Layout

```
legacy-personal-library/
├── slides/                     # slide assets: *.pptx (binaries not in repo)
│   └── README.txt
├── shapes/                     # shape assets: *.pptx + <name>.pptx.meta.json sidecars
│   ├── arrow-block.pptx.meta.json
│   ├── sample-note.png         # standalone PNG (listed by legacy My Assets *.png scan)
│   └── kpi/                    # nested subfolder is part of the legacy contract
│       └── kpi-card.pptx.meta.json
└── images/
    └── sample-logo.png
```

## Format facts the fixture encodes

- `.meta.json` sidecars are named `<asset-file>.meta.json` (e.g. `arrow-block.pptx.meta.json`
  for `arrow-block.pptx`) and contain a JSON array of `ShapeMetadataItem` serialized by
  `DataContractJsonSerializer` (PascalCase members):
  `OriginalShapeId`, `Name`, `Left`, `Top`, `Width`, `Height` (points at save time),
  `SlideWidth`, `SlideHeight` (source slide size in points), `Keywords`.
- Geometry values in the samples use a 960×540 (16:9) source slide.
- The sidecar can exist without its `.pptx` in this fixture; consumers must treat a
  missing binary as a corrupt/missing-asset case, not crash.
- Image extensions recognized by the legacy scan: `.png`, `.jpg`, `.jpeg`, `.svg`
  (manager scan) / plus `.gif`, `.bmp` (favorites scan).

## Deterministic file checksums (SHA-256)

| File | SHA-256 |
| --- | --- |
| `shapes/arrow-block.pptx.meta.json` | `f0ab1863a0b73e7fa46959df8058214f80cf890e5f53a0b46dcf50ddf1eed0b6` |
| `shapes/kpi/kpi-card.pptx.meta.json` | `532a4684ed27c661c699df4a4591e76e183519c6c5345fb3829d0ae657df9451` |
| `shapes/sample-note.png` | `497790947d4666760ce38f3c00e852c71fdb66cae849bae8e9ede352719e1581` |
| `images/sample-logo.png` | `497790947d4666760ce38f3c00e852c71fdb66cae849bae8e9ede352719e1581` |

## Producing the live fixture with PowerPoint

To produce the full live equivalent for real-host evidence (owner-executed):

1. Copy this folder tree to `%APPDATA%\oPenEfficiency\favorites\` (or run against the
   real personal library).
2. In PowerPoint, select slides/shapes and use Asset Manager → Save Selection to create
   real `.pptx` assets next to the sidecars above.
3. Record observations per `docs/testing/feat-008-current-asset-manager-compatibility.md` §8.

## Consumers

- `oPenEfficiency.Core.Tests` (AM-S1+ tests; folder is copied to output via the
  `Fixtures\**\*` content glob in `oPenEfficiency.Core.Tests.csproj`).
- Migration/index code must never modify this folder; it is read-only input.
