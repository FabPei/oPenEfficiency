# Chart fixture deck specification (CH-S0)

Spec for the owner-created chart fixture deck referenced by
`docs/testing/feat-006-current-charting-baseline.md` §5. The deck itself is
owner-created in PowerPoint (native chart binaries are not hand-authored in this
repo) and stored under the git-ignored evidence folder
`artifacts/feat-006-baseline/feat006-chart-fixture.pptx`; record its SHA-256 in the
results notes so the baseline stays reviewable.

## Required slides

1. **Single-series column chart** — 6–8 points, positive values (Delta/CAGR/Growth
   baseline case).
2. **Single-series column chart with a negative start value** — CAGR edge case
   (limitation L3).
3. **Stacked column chart** — 3 series × 6 points (stacked-total math intended case).
4. **Clustered column chart** — 3 series × 6 points (cross-series-sum limitation L2).
5. **100% stacked bar chart** — percent semantics.
6. **Line chart** — single series with one obvious spike (outlier preset).
7. **Combination chart** — column + line series (multi-series behavior).
8. **Pie chart** — unsupported-family behavior for level line (L5).
9. **Two charts on one slide** (column + line) — independent overlay targeting.
10. **One column chart inside a group shape** — group selection contract.
11. **One native linked chart** (linked to `chart-source.xlsx` in the same folder) —
    refresh behavior (FEAT-005 contract tie-in).
12. **Indicator slide** — Harvey Ball, Traffic Light, Thermometer, Star Rating,
    Progress Series, Checkbox (unaffected-by-overlay evidence).

## Conventions

- Keep default chart geometry (no manual moves) on slides 1–8 so overlay drift (L1)
  is attributable to the overlay being static, not to pre-moved charts.
- Name each chart shape via the selection pane (`FixtureColumn`, `FixtureStacked`, …)
  for deterministic notes.
- After executing the §5 checklist, append per-slide observations as
  `notes-<slide>.md` next to the deck and record the deck's SHA-256 in `notes-deck.md`.
