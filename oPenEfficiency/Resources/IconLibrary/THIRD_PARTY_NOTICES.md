# Third-party icon notices

oPenEfficiency redistributes SVG assets from the collections below. The
corresponding license texts are included in the `LICENSES` directory next to this
notice and must remain with redistributed icon assets.

## Bundled snapshot

The current manifest was generated on 2026-07-12 and committed to this repository
on 2026-07-24. The generator used local upstream checkouts but did not record
their commit hashes or release tags. The exact upstream revisions are therefore
unknown and must not be inferred from current upstream releases.

| Collection | Bundled SVGs | Source | License | Bundled revision |
|---|---:|---|---|---|
| Phosphor Icons | 4,293 | https://github.com/phosphor-icons/core | MIT | Unknown snapshot |
| Tabler Icons | 5,731 | https://github.com/tabler/tabler-icons | MIT | Unknown snapshot |
| Lucide Icons | 1,739 | https://github.com/lucide-icons/lucide | ISC; selected Feather-derived icons remain MIT | Unknown snapshot |
| Iconoir | 1,606 | https://github.com/iconoir-icons/iconoir | MIT | Unknown snapshot |
| Heroicons | 648 | https://github.com/tailwindlabs/heroicons | MIT | Unknown snapshot |

License status and upstream links were rechecked against the official repositories
on 2026-08-04. The generator now rejects dirty upstream checkouts and records the
exact Git commit in each collection's `SourceRevision` field before replacing the
bundled snapshot.

## Repository-specific custom icons

The `custom` collection currently contains 65 SVG assets generated with an AI
image-generation system at the direction of the project maintainer. The specific
model, prompts, and per-file generation dates were not retained, and the assets
are not represented as belonging to one of the external collections above. To
the extent that copyrightable human selection, arrangement, or modification is
present and controlled by the project maintainer, those contributions are
offered under the recorded MIT license (bundled as Resources/IconLibrary/LICENSES/Project-Assets-MIT.txt); no exclusive copyright is claimed in
purely AI-generated elements. Do not add externally sourced custom icons without
documenting author, source URL, source revision or creation date, and license in
the custom collection's `README.md`.

## Standalone application icons

The ribbon and application icons in `UI/Assets/Icons` are outside the searchable
icon-library manifest. Their Lucide, Phosphor, and project-specific provenance is
documented in the repository-level `THIRD_PARTY_NOTICES.md`, which is also
included in release payloads. The license texts in this directory apply to those
Lucide and Phosphor derivatives as well as to the searchable icon library.

## Copyright notices

- Phosphor Icons: Copyright (c) 2023 Phosphor Icons.
- Tabler Icons: Copyright (c) 2020-2026 Paweł Kuna.
- Lucide Icons: Copyright (c) 2026 Lucide Icons and Contributors.
- Feather-derived Lucide icons: Copyright (c) 2013-present Cole Bemis.
- Iconoir: Copyright (c) 2021 Luca Burgio.
- Heroicons: Copyright (c) Tailwind Labs, Inc.

The upstream projects and authors do not endorse oPenEfficiency. Names and links
are provided solely for attribution and license compliance.
