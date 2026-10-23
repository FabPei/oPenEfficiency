# From a fresh clone to an AI-assisted feature pull request

Use this guide for one small PowerPoint feature or correction. It works with the public
source tree and does not require private chat history, private agent files or a sibling
repository. The [feature request template](templates/ai-feature-request.md) is the brief to
fill in and paste into your coding AI. Product labels and help text must be English; your
discussion with the AI can use your preferred language.

## 1. Choose one result and prepare the machine

Search existing commands and issues first. Describe one observable result, what stays out
of scope and how a reviewer can tell it works. Avoid combining unrelated features in one PR.

Use Windows, desktop PowerPoint 2016 or later, Visual Studio with Office/VSTO development
tools and MSBuild, and the .NET Framework 4.8 SDK/targeting pack. Visual Studio/MSBuild must
support `oPenEfficiency.slnx`. Install Windows PowerShell 5.1, PowerShell 7 (`pwsh` on PATH),
Git and a .NET SDK matching the checked-out CI configuration (public CI uses `9.0.x` at
the time this guide was prepared). The .NET SDK does not replace net48 or Office tools.
GitHub CLI (`gh`) is needed only for the CLI fork/PR route below; a GitHub account is
needed to publish a contribution, not to read or clone public code.

Run commands from the repository root in PowerShell. Check each result before continuing;
stop on an error. Save verbose output locally and report failures honestly.

## 2. Fork, clone and create a branch

External contributors: run `gh auth status`; if needed, sign in using
`gh auth login --hostname github.com --git-protocol https --web` and configure Git with
`gh auth setup-git`. Complete sign-in yourself; do not paste tokens into the AI brief.
Then, from a parent folder in which `oPenEfficiency` does not already exist:

```powershell
gh repo fork FabPei/oPenEfficiency --clone=false --remote=false
$contributorLogin = gh api user --jq .login
git clone "https://github.com/$contributorLogin/oPenEfficiency.git"
Set-Location oPenEfficiency
git remote add upstream https://github.com/FabPei/oPenEfficiency.git
git fetch upstream main
git switch -c codex/my-feature upstream/main
git remote -v
git status --short --branch
```

If you already have a fork, clone that fork and fetch upstream; do not recreate it.
`origin` must be your fork; `upstream` is `FabPei/oPenEfficiency`. Replace `my-feature`
with a short task name. Fork settings must allow the intended branch and PR operation.

Maintainers who own the public repository cannot fork it to the same account. Instead,
clone `https://github.com/FabPei/oPenEfficiency.git` into a new folder and create
`codex/my-feature` from `origin/main`. Keep any public feature branch on public Git history.
A private development checkout follows its own repository instructions; never merge its
history into a public branch or copy its complete file tree.

## 3. Give the AI the brief

Open the cloned folder in your coding tool. Fill in the template and start with:

> Read `docs/development/ai-feature-contribution.md` and applicable repository instructions.
> Implement the filled brief below. Inspect the existing behavior first. Identify the exact
> files, UI family and checks before editing, then carry out the authorized work. Preserve
> unrelated edits and tell me clearly what remains unverified.

An AI may not automatically discover arbitrary Markdown files. Pass these paths explicitly.
Read a checkout's `AGENTS.md` when present; the public workflow here does not depend on one.
The filled publication field controls whether this task only prepares a patch or also pushes
the specified branch and opens a draft PR. Already granted authorization need not be asked again.

## 4. Find the correct extension points

Use [feature development](features.md) and inspect a nearby real implementation. This map
uses paths present in the public source; it does not imply that every feature needs every file.

| Responsibility | Start here |
| --- | --- |
| Host-independent calculations and deterministic MSTest cases | `oPenEfficiency.Core/`, `oPenEfficiency.Core.Tests/` |
| PowerPoint operation and selection | `oPenEfficiency/Features/`, `oPenEfficiency/Services/PowerPointManager.cs` |
| Existing small alignment action | `oPenEfficiency/Features/Alignment/MatchWidthToFirstFeature.cs`, `AlignmentHelpers.cs` in the same folder |
| Attribute fields and resolved catalog policy | `oPenEfficiency.Core/Services/Attributes/FeatureMetadataAttribute.cs`, `oPenEfficiency.Core/UI/FeatureCatalogPolicy.cs` |
| Discovery, help and default layout | `oPenEfficiency/UI/FeatureLibrary.cs` |
| Typed execution and generated registrations | `oPenEfficiency/UI/Panels/MainSidebar.Commands.cs`, `MainSidebar.GeneratedCommands.cs` in the same folder; `eng/GenerateCommandRegistrations.ps1` |
| Existing WPF styles and theme resources | `oPenEfficiency/UI/Styles/Controls.xaml` |
| Optional Office ribbon exposure | `oPenEfficiency/UI/MainRibbon.xml`, callbacks in `oPenEfficiency/ThisAddIn.cs` |
| User documentation | `README.md`, `docs/Features.md`, `docs/product/feature-catalog.md` |
| Contract and integration-boundary tests | `oPenEfficiency.Catalog.Tests/`, `oPenEfficiency.Command.Tests/`; inspect their actual public fixtures |

### Feature rules

- Prefer complete `[FeatureMetadata]` with a unique stable ID, English name, description,
  practical detailed help and accurate selection requirements. IDs are persistence keys.
- The resolved catalog is the source for Sidebar, Settings search, Feature Explorer and
  applicable ribbon/context surfaces. Check every relevant surface; explain intentional omissions.
  Set category, maturity and placement in the existing policy. Do not invent another command list.
- Straightforward actions use generated typed registrations. Regenerate with
  `./eng/GenerateCommandRegistrations.ps1`; use explicit registrations for arguments,
  toggles or dialogs. Never add a reflection-based execution fallback.
- Put pure calculations in Core and keep Office/WPF references out of Core. Use existing
  `PowerPointManager` wrappers for COM, respect 1-based collections, selection scope and
  object ownership. Do not release shared host objects or change unrelated presentation state.
- Define Undo, empty/invalid selections, unsupported objects, cancellation and partial-failure
  behavior before writing mutations. Log caught failures with context using the existing
  `ExceptionLogger`; do not silently claim success after a failed operation.
- Use existing shape-tag/`AlternativeText` persistence patterns and `oPE_` metadata prefixes
  when persistence is needed. A one-shot action need not invent stored state.
- Keep names, discovery, tooltips, help and user documentation consistent. A tooltip should add
  useful behavior information rather than create another feature name.

### Choose the UI family before creating UI

| User need | UI family | Starting contract |
| --- | --- | --- |
| Immediate action on the current selection | Existing command button | Reuse the existing Sidebar/ribbon command surface; no new window required |
| A few contextual options after a click/right-click | Context Popup / On-click Control | Small content-sized popup; safe outside-click/Esc dismissal |
| Repeated adjustments while editing a slide | Floating Tool / Palette | Compact modeless tool owned by the actual PowerPoint window |
| Related inputs for one bounded operation | Compact Task Window | Compact form; explicit Apply/Cancel and modal/modeless choice |
| Continuous search, libraries or commands | Sidebar / Task Pane | Persistent grouped controls following the host lifecycle |
| Broad browsing or configuration | Large Workspace | Navigation and progressive disclosure; use only when the task needs the space |

Reuse the corresponding shell/styles in `Controls.xaml`, semantic `DynamicResource`
brushes and existing spacing/typography tokens. Inherit the application font. Keep layouts
compact, group related controls, and avoid repeated headings, unnecessary cards and empty space.
Support Light/Dark, keyboard/visible focus, accessible control names, text growth and DPI.
Floating surfaces need the actual Office HWND as owner, proper cleanup, focus restoration and
monitor-clamped DPI-aware placement. `Topmost` is not a replacement for ownership.

**Icons today:** reuse a suitable approved existing icon unchanged, recording its source.
The minimal standalone icon Composer is not bundled in the public clone yet. Creation or
redrawing of project icons still needs the maintainer's approved Composer/export process.
Do not hand-draw replacement SVG/path geometry or pretend a private sibling tool is available.
Request the missing authoring input if the task cannot use an existing icon. This limitation
does not prevent code and documentation work that needs no new icon geometry.

## 5. Check the work at the right level

Before a first **code** change, establish a baseline with
`./eng/Build.ps1 -Configuration Release` and retain its result. The public build creates
temporary local VSTO manifest-signing material and restores its local state; no maintainer
key is needed. Visual Studio/F5 deployment is a separate operation and can alter add-in
registration. Use an isolated development environment for that operation.

During implementation, use the cheapest relevant checks:

| Changed input | Fast feedback |
| --- | --- |
| Prose only | `git diff --check`, link/path/example checks; no application build solely for prose |
| Feature signature/registration | `./eng/GenerateCommandRegistrations.ps1 -Check` after regeneration |
| Core/host boundaries | `./eng/Test-Architecture.ps1` |
| WPF/styles | `./eng/Test-VisualConsistency.ps1` plus the actual relevant UI observations |
| Pure logic/catalog/command contracts | Build affected dependencies, then run the appropriate MSTest project/filter |

Example for an existing affected Core test (replace the filter with the task's actual test):

```powershell
dotnet test oPenEfficiency.Core.Tests/oPenEfficiency.Core.Tests.csproj --configuration Release --filter "FullyQualifiedName~HarveyBallStatePolicyTests" --logger "trx;LogFileName=targeted-core.trx"
```

Do not use `--no-build` on an unbuilt or stale test assembly. Public contract fixtures are
editable source inputs: update them only for an intentional documented contract change,
never to conceal a failure. Add deterministic cases for the new behavior; reproduce a bug
before fixing it where feasible. Separate prepared cases from actually executed results.

If `eng/Test-Changed.ps1` exists in your checkout, use it with exact affected paths and inspect
its plan. **It is absent from the current public source**; do not copy commands that depend on
private tools. The public checks listed above remain available.

At the final **code/test/asset/build-infrastructure** milestone, after independent review and
corrections, run `./eng/Build.ps1 -Configuration Release -Installer` once. Retain full logs and
TRX results; repeat only after an affecting change or to verify a fix for a failed run.
The private repository retains its own canonical checkpoint rules. A full build does not
prove PowerPoint behavior, installation or complete product acceptance.
Give the final checkpoint evidence to the independent reviewer before marking the work verified.

For affected Office behavior, check relevant real-host cases in an isolated Office session
when feasible: normal/empty/unsupported selection, Undo, save/reopen when state persists,
Light/Dark and relevant keyboard/DPI behavior. Never attach automation to a user's live
session, install an MSI or reset registration as a shortcut. Preserve running Office and
restore temporary registry state. If the environment is unavailable, record the missing
evidence and keep the PR as a draft; do not label it Passed.

## 6. Example brief: one new alignment action

This is a **design example**, not an implemented or shipped command. It shows how to fill the
template without adding a new dialog or icon tool dependency:

| Field | Example |
| --- | --- |
| Result | A command named **Average Width** sets selected objects to the arithmetic mean of their original widths. First search the catalog for equivalent behavior. |
| Scope | At least two supported selected shapes on the active slide; only width changes. No ribbon button, popup, settings, new icon geometry or saved state. |
| Selection/failure | Disable for no/one shape or wrong selection type. Define supported objects before implementation. Validate all widths first; reject non-finite/non-positive values without mutation. Define how COM failure is reported and how Undo recovers changes. |
| Geometry | Preserve heights, positions and each object's aspect-ratio lock state. Inspect the existing Match Width helper before choosing the host mutation path. |
| Paths to propose | New `oPenEfficiency.Core/Services/AverageWidthPolicy.cs`, `oPenEfficiency.Core.Tests/AverageWidthPolicyTests.cs`, `oPenEfficiency/Features/Alignment/AverageWidthFeature.cs`; existing catalog policy, discovery/default layout, generated registry and user docs as required. These new files do not exist yet. |
| UI/icon | Existing command-button family; reuse the suitable Match Width icon unchanged after checking its source. English help explains selection and width-only behavior. |
| Checks | Pure cases: widths 10/20/30 yield 20; equal widths remain equal; invalid/empty/single input rejected; inputs remain unchanged. Host cases: selection scope, width-only changes, lock-state preservation, Undo, no-selection state and truthful COM error handling. |
| Review | Independently verify catalog/discovery/typed dispatch, deterministic cases, host ownership and documentation; retain unrun host cases explicitly. |

Inspect `MatchWidthToFirstFeature.cs` and its helper as the existing extension pattern; do not
duplicate **Match Width**, which already exists. A contributor can supply their own equally
bounded brief instead. This guide has not executed or validated the hypothetical feature.

## 7. Review, commit and prepare the PR

Have a separate reviewer inspect the final diff and evidence. A second AI context or another
person can review; the implementer cannot approve itself. Resolve findings and rerun only
affected checks. Record unavailable independent review or host evidence as pending. Keep the
summary brief: what changed, why, checks actually performed and remaining limitations; no
mandatory detailed numerical report. Detailed evidence belongs in files or CI artifacts.

Review `git diff --name-only` and stage only the task's exact files with `git add -- ...`
(replace `...` with reviewed paths). Then:

```powershell
git diff --cached --check
git diff --cached --stat
git commit -m "feat: describe the implemented result"
```

Copy [the PR template](../../.github/PULL_REQUEST_TEMPLATE.md) to a local file such as
`TestResults/pr-body.md`, fill it in and review it before sending. Include clean evidence
attachments only when useful; never include private keys, credentials or confidential slides.

When publishing this exact target has been authorized in the brief, use:

```powershell
$contributorLogin = gh api user --jq .login
$featureBranch = git branch --show-current
gh pr list --repo FabPei/oPenEfficiency --head $featureBranch --json number,url,headRepositoryOwner,headRefName
git remote get-url origin
# Confirm origin is the intended public repository/fork and the diff is public-safe.
git push --set-upstream origin $featureBranch
gh pr create --repo FabPei/oPenEfficiency --base main --head "${contributorLogin}:$featureBranch" --draft --title "Describe the implemented result" --body-file TestResults/pr-body.md
```

Review the returned PR URL and use `gh pr checks <PR-URL> --repo FabPei/oPenEfficiency`
with that actual URL. A pending/missing check is not success. If a matching PR already exists,
update that branch instead of opening a duplicate. Confirm the actual fork owner when it
differs from the signed-in account, including for matching PRs returned by the branch filter.
`gh pr list --head` takes a branch name; the owner-qualified form is used by `gh pr create`
instead. Push/PR errors require checking remote state before retrying.
Without GitHub permission, retain the local commit/patch and the intended PR body and report
the blocker. Do not claim a PR exists until GitHub returns and confirms it.

The maintainer decides merge and release. A contributor PR does not authorize a tag, merge,
release, installer installation, change of license or public transfer of private files. MSI
releases remain unsigned; technically required VSTO manifest signatures are separate.

## References

- [GitHub CLI fork options](https://cli.github.com/manual/gh_repo_fork),
  [login](https://cli.github.com/manual/gh_auth_login) and
  [PR creation](https://cli.github.com/manual/gh_pr_create): explicit remote/head/base, draft and body-file options.
- [PR listing](https://cli.github.com/manual/gh_pr_list): use the branch-only head filter and
  confirm the returned PR belongs to the intended fork before reusing it.
- [Public development wiki](https://github.com/FabPei/oPenEfficiency/wiki/Development):
  public setup, VSTO development and source layout.
- [Contribution entry point](../../CONTRIBUTING.md) and [project license](../../LICENSE).
