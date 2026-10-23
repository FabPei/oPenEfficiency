# AI feature request — copy and fill in

Replace every bracketed field. Keep one result per task. Do not include credentials or
confidential presentation content. Tell the AI to read
`docs/development/ai-feature-contribution.md` and applicable repository instructions first.

## Result and boundaries

- User problem: [Who needs to do what?]
- Observable result: [One sentence describing the desired behavior.]
- Existing behavior/nearest feature: [Name, stable ID and source paths, or ask the AI to locate them.]
- Work type: [new command / correction / UI adjustment / documentation]
- In scope: [Exact behavior and surfaces.]
- Out of scope: [Related ideas deliberately excluded.]

## Selection and operation

- Supported host and scope: [PowerPoint version/bitness, active slide/selected shapes/etc.]
- Supported objects: [Text shapes, tables, groups, charts, etc.; make exclusions explicit.]
- Selection rule: [Minimum/maximum, anchor/order, behavior after the selection changes.]
- Normal result: [Before/after example.]
- Empty/unsupported/invalid input: [Disabled command, validation message, no-op, etc.]
- Cancellation/partial failure: [What remains unchanged? How are errors logged and explained?]
- Undo: [Expected restoration behavior; explain if this is read-only.]
- Persistence: [None, existing shape tags/AlternativeText/configuration; save/reopen expectation.]
- Must preserve: [Other shapes, geometry/formatting, aspect locks, user selection and document state.]

## UI and discovery

- UI family: [existing command / context popup / floating palette / compact task window /
  sidebar / large workspace; justify any new window]
- English command name and proposed stable ID: [Name, ID; check existing IDs before choosing.]
- Help/tooltip: [Practical usage; avoid repeating the name without adding information.]
- Surfaces: [Sidebar, Settings search, Feature Explorer, default placement, optional ribbon/context.]
- Shared styles/theme: [Existing neighboring surface and Controls.xaml resources.]
- Keyboard/focus/DPI/text growth: [Applicable behavior and checks.]
- Icon: [Suitable existing approved icon and source; no new geometry unless the approved
  Composer/export path and its inputs are actually available.]

## Scope of files and checks

- Existing files to inspect: [Exact paths or one bounded search for the nearest implementation.]
- Expected new files: [Core policy, host adapter, deterministic tests, docs, or none.]
- Deterministic examples: [Input → result, plus invalid/boundary cases.]
- Host/UI checks: [Relevant real PowerPoint/Undo/persistence/theme cases.]
- Available environment: [Windows/Office/tool versions; can an isolated Office session be used?]
- Missing evidence: [Known unavailable machine, icon tool or owner observation.]
- Review criteria: [Observable conditions a separate reviewer must verify.]

## Git and publication authorization

- Repository and base: [Exact URL/local path and base branch or tag.]
- Working branch: [codex/short-task-name]
- Authorized actions: [prepare local patch only / commit / push specified branch / open draft PR]
- Push target and PR base, if authorized: [Exact repository/fork owner and branch; e.g.
  your public fork → FabPei/oPenEfficiency main.]
- Excluded actions: [No merge, tags, release or unrelated/private file transfer.]

## Instructions to the coding AI

1. Read applicable repository instructions and the guide. Start with status and a bounded
   inspection; preserve unrelated edits. Confirm the requested behavior is not already present.
2. Briefly identify the relevant files, UI family and checks. Resolve material ambiguity
   before dependent edits; continue independent work. Do not invent missing private tooling.
3. Implement only this brief using existing catalog/typed dispatch, shared UI and COM patterns.
   Keep pure logic in Core, English product copy and documentation/discovery consistent.
4. Add relevant deterministic regression/feature cases for code changes and use targeted
   checks during development. Follow the guide's final code checkpoint and applicable
   repository validation rules. Prose alone does not require a product build.
5. Preserve running Office and registry state. Use isolated host checks when feasible and
   record unrun evidence honestly. Never weaken a test or call a planned case Passed.
6. Obtain independent review of product/instruction-policy changes; address findings before
   the final checkpoint. If unavailable, leave that requirement explicitly pending.
7. Review exact staged files. Perform only the commit/push/PR actions authorized above and
   applicable repository instructions; verify the target before any public write. Reuse
   existing authorization rather than asking again. Never expose private history or keys.
8. Finish with a short explanation of the result, evidence links, remaining limitations and
   one next step. No mandatory detailed check counts. Report an actual PR URL only after
   verifying that the PR exists.
