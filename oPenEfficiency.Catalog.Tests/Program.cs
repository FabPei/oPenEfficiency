using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Services.Shortcuts;
using oPenEfficiency.UI;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestAssert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;

namespace oPenEfficiency.Catalog.Tests
{
    [TestClass]
    public class Program
    {
        private static readonly RegexOptions SourceOptions =
            RegexOptions.Singleline | RegexOptions.CultureInvariant;

        private static void WithCatalog(Action<string, IReadOnlyList<FeatureCatalogRecord>> assertion)
        {
            string root = FindRepositoryRoot();
            assertion(root, ReadCurrentCatalog(root));
        }

        [TestMethod] public void CurrentCatalogIsValid() => WithCatalog((root, catalog) => AssertNoErrors("current catalog", FeatureCatalogValidator.ValidateCatalog(catalog)));
        [TestMethod] public void CanonicalFeatureInventoryIsCurrent() => WithCatalog(ValidateCanonicalFeatureInventory);
        [TestMethod] public void CurrentPresetsAreValid() => WithCatalog((root, catalog) => ValidateCurrentPresets(catalog));
        [TestMethod] public void CatalogRulesHaveCoverage() => ValidateRuleCoverage();
        [TestMethod] public void CommandsMatchCatalog() => WithCatalog(ValidateCommandCoverage);
        [TestMethod] public void SidebarSelectionLogicIsValid() => ValidateSidebarUiLogic();
        [TestMethod] public void Wave1DistributeHorizontalSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnDistributeHorizontal"));
        [TestMethod] public void Wave1DistributeVerticalSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnDistributeVertical"));
        [TestMethod] public void Wave1DistributeHorizontalSelectionGuardIsTheProductionRule() => WithCatalog((root, catalog) => ValidateDistributeSelectionGuardDelegation(root, catalog, "BtnDistributeHorizontal"));
        [TestMethod] public void Wave1DistributeVerticalSelectionGuardIsTheProductionRule() => WithCatalog((root, catalog) => ValidateDistributeSelectionGuardDelegation(root, catalog, "BtnDistributeVertical"));
        [TestMethod] public void Wave1DistributeContextMenuRouteMatchesItsSelectionScope() => WithCatalog((root, catalog) => ValidateDistributeContextMenuRoutes(root));
        [TestMethod] public void ToggleGalleryOnlyTextContractIsValid() => WithCatalog((root, catalog) => ValidateToggleGalleryOnlyTextContract(root, catalog));
        [TestMethod] public void SearchBarEnterPriorityContractIsValid() => ValidateSearchEnterPriorityContract();
        [TestMethod] public void QuickContextMenuDefaultsResolveInCatalogIsValid() => WithCatalog((root, catalog) => ValidateQuickContextMenuDefaults(catalog));

        /// <summary>
        /// REG-TEXTFONTPROPS-001 disposition (b) (reviewer, 2026-10-03): PowerPoint's Font dispatch has no
        /// StrikeThrough/Spacing member, so Strikethrough and Character Spacing document the
        /// toggle/gallery-only contract. Batch 12 (ca33eb0d, 917e75c6) removed the dead explicit paths and
        /// corrected the catalog variant vocabulary; a reviewer rejected the first closure precisely because
        /// the old variants kept re-entering the discovery text. This pins the production facts that fix:
        /// the catalog vocabulary, the discovery text, and the context-menu route each feature dispatches.
        /// </summary>
        private static void ValidateToggleGalleryOnlyTextContract(
            string repositoryRoot,
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            var expectations = new[]
            {
                new[]
                {
                    "BtnCharacterSpacing", "More Spacing...",
                    "CharacterSpacingFeature.ExecuteMoreSpacing(GetManager())",
                    "CharacterSpacingFeature.ExecuteSpecific", "Right-click to apply Very Tight"
                },
                new[]
                {
                    "BtnFormatStrikethrough", "Toggle Strikethrough Formatting",
                    "FormatStrikethroughFeature.Execute(GetManager())",
                    "FormatStrikethroughFeature.Apply", "FormatStrikethroughFeature.Remove"
                },
            };

            foreach (var expectation in expectations)
            {
                string id = expectation[0];
                string variant = expectation[1];
                string requiredRoute = expectation[2];

                var record = catalog.FirstOrDefault(candidate => candidate.Id == id);
                Assert(record != null, $"Catalog record {id} is missing.");
                if (record == null) continue;

                TestAssert.IsNotNull(record.Variants, $"{id}: the catalog exposes no variant labels.");
                TestAssert.AreEqual(1, record.Variants.Count,
                    $"{id}: expected exactly the corrected single variant, found [{string.Join(" | ", record.Variants)}].");
                TestAssert.AreEqual(variant, record.Variants[0],
                    $"{id}: the catalog variant vocabulary regressed to a dead label.");

                string help = record.DetailedHelpText ?? string.Empty;
                foreach (int deadIndex in new[] { 3, 4 })
                {
                    string deadPhrase = expectation[deadIndex];
                    if (string.IsNullOrEmpty(deadPhrase)) continue;
                    Assert(!help.Contains(deadPhrase),
                        $"{id}: discovery help still advertises the dead path '{deadPhrase}'.");
                }

                Assert(help.Contains(variant) || help.Contains(id == "BtnCharacterSpacing" ? "More Spacing" : "Right-click repeats the toggle"),
                    $"{id}: the corrected variant is neither part of the help text nor discoverable.");

                // Route pin: the context menu dispatches only the working native path.
                string menuSource = File.ReadAllText(Path.Combine(
                    repositoryRoot, "oPenEfficiency", "UI", "Panels", "MainSidebar.ContextMenus.TextFormatting.cs"));
                var block = Regex.Match(
                    menuSource,
                    @"else if \(btn\.Name == """ + Regex.Escape(id) + @"""(?<body>.*?)(?=\n\s*else if)",
                    SourceOptions);
                Assert(block.Success, $"MainSidebar.ContextMenus.TextFormatting.cs: the {id} menu block was not found.");
                if (!block.Success) continue;
                string body = block.Groups["body"].Value;
                TestAssert.AreEqual(1, CountOccurrences(body, "GetCatalogVariantLabel"),
                    $"{id}: the context menu must expose exactly one corrected variant item.");
                Assert(body.Contains(requiredRoute),
                    $"{id}: the context menu no longer routes to {requiredRoute}.");
                Assert(!body.Contains(expectation[3]) && !body.Contains(expectation[4]),
                    $"{id}: the context menu resurrected a dead explicit-action route.");
            }
        }

        /// <summary>
        /// Review round 6h measured that swapping the relativeToSlide value at the context-menu call sites inverted which
        /// route every variant applies (and the labels kept their meaning) while every gate script, both test legs and the
        /// acceptance driver stayed green - the caller is part of the contract. This pins the pair of the catalog variant
        /// label and the route the variant dispatches.
        /// </summary>
        private static void ValidateDistributeContextMenuRoutes(string repositoryRoot)
        {
            var path = Path.Combine(repositoryRoot, "oPenEfficiency", "UI", "Panels", "MainSidebar.ContextMenus.TextFormatting.cs");
            var source = File.ReadAllText(path);
            var sites = 0;
            foreach (var entry in new[]
            {
                new[] { "BtnDistributeHorizontal", "DistributeHorizontalFeature" },
                new[] { "BtnDistributeVertical", "DistributeVerticalFeature" },
            })
            {
                var block = Regex.Match(source, @"btn\.Name == """ + entry[0] + @"""(?<body>.*?)(?=\n\s*else if|\n\s*\}\s*\n\s*\})", RegexOptions.Singleline);
                Assert(block.Success, $"MainSidebar.ContextMenus.TextFormatting.cs: the {entry[0]} menu block was not found.");
                if (!block.Success) continue;

                var body = block.Groups["body"].Value;
                var rows = Regex.Matches(body, @"Header = GetCatalogVariantLabel\(btn\.Name, \d+, ""(?<label>[^""]+)""\)[^\n]*\n\s*[^\n]*?" + entry[1] + @"\.Execute\(GetManager\(\), (?<args>[^)]*)\)");
                Assert(rows.Count == 4, $"MainSidebar.ContextMenus.TextFormatting.cs: {entry[0]} must offer its four menu variants (found {rows.Count}).");
                foreach (Match row in rows)
                {
                    var label = row.Groups["label"].Value;
                    var args = row.Groups["args"].Value;
                    sites++;
                    var route = Regex.IsMatch(args, @"relativeToSlide:\s*true") ? true
                        : Regex.IsMatch(args, @"relativeToSlide:\s*false") ? false
                        : (bool?)null;
                    Assert(route.HasValue, $"MainSidebar.ContextMenus.TextFormatting.cs: the '{label}' variant does not state which route it applies.");
                    if (!route.HasValue) continue;

                    var acrossSlide = label.IndexOf("Across Slide", StringComparison.Ordinal) >= 0;
                    var withinSelection = label.IndexOf("Within Selection", StringComparison.Ordinal) >= 0;
                    Assert(acrossSlide || withinSelection, $"MainSidebar.ContextMenus.TextFormatting.cs: the '{label}' variant names neither route, so its label and the route it applies cannot be checked as a pair (review round 6h).");
                    Assert(acrossSlide != withinSelection, $"MainSidebar.ContextMenus.TextFormatting.cs: the '{label}' variant names the across-slide and the within-selection route at once.");
                    Assert(route.Value == acrossSlide, $"MainSidebar.ContextMenus.TextFormatting.cs: the '{label}' variant dispatches relativeToSlide: {route.Value.ToString().ToLowerInvariant()} - the label and the route it applies must stay a pair (review round 6h).");
                }
            }

            Assert(sites == 8, $"MainSidebar.ContextMenus.TextFormatting.cs: expected eight distribute context-menu call sites, found {sites}.");
        }
        [TestMethod] public void Wave1RotateShapeSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnRotateShape"));
        [TestMethod] public void Wave1RectifyRotationSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnRectifyRotation"));
        [TestMethod] public void Wave1FlipHorizontalSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnFlipHorizontal"));
        [TestMethod] public void Wave1FlipVerticalSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnFlipVertical"));
        [TestMethod] public void Wave1MatchSizeSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnMatchSize"));
        [TestMethod] public void Wave1LockDimensionsSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnLockDimensions"));
        [TestMethod] public void Wave1SelectSameTypeSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnSelectSameType"));
        [TestMethod] public void Wave1HideSelectedSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnHideSelected"));
        [TestMethod] public void Wave1SlideSizeSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnSlideSize"));
        [TestMethod] public void Wave1MatchHeightSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnMatchHeight"));
        [TestMethod] public void Wave1MatchWidthSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnMatchWidth"));
        [TestMethod] public void Wave1AdjustSpacingSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnAdjustSpacing"));
        [TestMethod] public void Wave1RemoveHorizontalGapsSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnRemoveHorizontalGaps"));
        [TestMethod] public void Wave1RemoveVerticalGapsSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnRemoveVerticalGaps"));
        [TestMethod] public void Wave1ShowHiddenSelectionScopeIsValid() => WithCatalog((root, catalog) => ValidateWave1SelectionScopeContract(catalog, "BtnShowHidden"));
        [TestMethod] public void LifecycleArchitectureIsValid() => WithCatalog((root, catalog) => ValidateLifecycleArchitecture(root));
        [TestMethod] public void CanonicalAliasesAreValid() => WithCatalog((root, catalog) => ValidateCanonicalAliases(catalog));
        [TestMethod] public void NamingAndBehaviorConsistencyIsValid() => WithCatalog(ValidateNamingAndBehaviorConsistency);
        [TestMethod] public void SavedLayoutNamesMigrate() => ValidateSavedLayoutNameMigration();
        [TestMethod] public void PresetGenerationPreservesCustomization() => WithCatalog((root, catalog) => ValidatePresetGenerationPreservesCustomizedLayout(catalog));
        [TestMethod] public void ShortcutSafetyIsValid() => WithCatalog((root, catalog) => ValidateShortcutSafety(root));
        [TestMethod] public void RibbonAndDocumentationMatchCatalog() => WithCatalog(ValidateRibbonAndDocumentationTruth);
        [TestMethod] public void IconAccentAuditMatchesCommandCatalog() => WithCatalog(ValidateIconAccentAudit);
        [TestMethod] public void ReviewerIconAccentContractMatchesCommandAssets() => WithCatalog(ValidateReviewerIconAccentContract);
        [TestMethod] public void GanttRulesModeCanBeReenabled() => ValidateGanttRulesModeToggle();
        [TestMethod] public void ChartOverlaysSupportGroupedCharts() => ValidateGroupedChartOverlays();
        [TestMethod] public void ProgressSeriesReadsVisibleHighlightColor() => ValidateProgressSeriesColorDetection();
        [TestMethod] public void CleanerCoversSafeSanitizationActions() => ValidateCleanerSanitizationActions();
        [TestMethod] public void ConnectedTextLabelClientAndSharedContractsAreValid() => ValidateConnectedTextLabels();
        [TestMethod] public void SidebarThemeTrackingContractIsValid() => ValidateSidebarThemeTrackingContract();
        [TestMethod] public void SidebarModelessWindowLifecycleIsValid() => ValidateSidebarModelessWindowLifecycle();
        [TestMethod] public void SidebarManagerFailureIsVisible() => ValidateSidebarManagerFailurePath();
        [TestMethod] public void SidebarNavigationFailuresPropagate() => ValidateSidebarNavigationFailurePath();
        [TestMethod] public void RegisteredSidebarDialogFailuresPropagate() => ValidateRegisteredSidebarDialogFailurePaths();
        [TestMethod] public void WpfDispatcherFailuresAreVisible() => ValidateWpfDispatcherFailureFallback();
        [TestMethod] public void MouseHookCallbacksAreComFree() => ValidateMouseHookCallbacksAreComFree();
        [TestMethod] public void SelectionDebounceDoesNotRetainTransientComObjects() => ValidateSelectionDebounceDoesNotRetainTransientComObjects();
        [TestMethod] public void SelectionMetadataTreatsMissingWindowAsExpectedComProbe() => ValidateSelectionMetadataWindowProbe();
        [TestMethod] public void SidebarPaneCleanupTreatsDisposedCollectionAsExpected() => ValidateSidebarPaneCleanupDisposalBoundary();
        [TestMethod] public void AlignAnchorOracleFollowsTheConfiguredAnchorMode() => ValidateAlignAnchorOracleConfiguration();
        [TestMethod] public void BuildStampIdentityComparisonNormalizesVersionComponents() => ValidateBuildStampIdentityComparison();
        [TestMethod] public void InstallerIdentityModeReachesNestedBuilds() => ValidateInstallerIdentityModeHandOff();
        [TestMethod] public void IllustrationEditableInsertionKeepsTheSourceName() => ValidateIllustrationEditableInsertionNaming();
        [TestMethod] public void AcceptanceRunnerArgumentsBindByName() => ValidateAcceptanceRunnerArgumentBinding();
        [TestMethod] public void FeatureIconGlyphsStayNonHitTestable() => ValidateFeatureIconGlyphHitTesting();
        [TestMethod] public void CopyPasteXyIsStrictlyPositionOnly() => ValidateCopyPasteXyPositionOnly();
        [TestMethod] public void QuickContextMenuConfigurationIsBoundedAndStable() => ValidateQuickContextMenuConfiguration();
        [TestMethod] public void ShapeFillPaletteMathIsStable() => ValidateShapeFillPaletteMath();
        [TestMethod] public void PerformanceModeIdentitySelectionIsSafe() => ValidatePerformanceModeIdentitySelection();
        [TestMethod] public void CountryFlagAssetsAreSafeAndComplete() => ValidateCountryFlagAssets();
        [TestMethod] public void IllustrationAssetsAreSafeAndComplete() => ValidateIllustrationAssets();
        [TestMethod] public void ConnectedTextPayloadMatchesLegacyContract() => ValidateConnectedTextPayloadContract();

        private static void ValidateConnectedTextPayloadContract()
        {
            // HTML builder: legacy marker spelling + encoding + raw <br /> inserted after encoding.
            string html = oPenEfficiency.OneNote.ConnectedTextPayload.BuildHtml("abc123", "A<B\r\nC&D");
            TestAssert.AreEqual(
                "<span data-ope-source=\"abc123\">A&lt;B<br />C&amp;D</span>",
                html,
                "Connected-text HTML payload must keep the exact legacy shape.");
            TestAssert.AreEqual("data-ope-source=\"abc123\"", oPenEfficiency.OneNote.ConnectedTextPayload.SourceMarker("abc123"));
            TestAssert.IsTrue(html.Contains(oPenEfficiency.OneNote.ConnectedTextPayload.SourceMarker("abc123")),
                "Built HTML must contain the source marker used as the refresh fallback.");

            // Page lastModifiedTime extraction (expected-last-modified concurrency token).
            var page = System.Xml.Linq.XDocument.Parse(
                "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" " +
                "ID=\"{1}\" lastModifiedTime=\"2026-09-18T10:00:00.000Z\" />");
            TestAssert.AreEqual(
                new DateTime(2026, 9, 18, 10, 0, 0, DateTimeKind.Utc),
                oPenEfficiency.OneNote.ConnectedTextPayload.ReadLastModified(page).ToUniversalTime(),
                "lastModifiedTime must parse as the update concurrency token.");
            TestAssert.AreEqual(DateTime.MinValue,
                oPenEfficiency.OneNote.ConnectedTextPayload.ReadLastModified(
                    System.Xml.Linq.XDocument.Parse("<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" />")),
                "A page without lastModifiedTime must yield MinValue.");

            // Record shape: legacy connected_text_labels.json property names must survive.
            var record = new oPenEfficiency.OneNote.ConnectedTextTargetRecord
            {
                SourceId = "src1", Host = "OneNote", DocumentPath = "{page-id}", TargetId = "src1"
            };
            var serializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(
                typeof(oPenEfficiency.OneNote.ConnectedTextTargetRecord));
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, record);
                string json = Encoding.UTF8.GetString(stream.ToArray());
                foreach (string legacyProperty in new[] { "SourceId", "Host", "DocumentPath", "TargetId" })
                    TestAssert.IsTrue(json.Contains("\"" + legacyProperty + "\":"),
                        "Legacy record property missing from JSON: " + legacyProperty);

                stream.Position = 0;
                var restored = (oPenEfficiency.OneNote.ConnectedTextTargetRecord)serializer.ReadObject(stream);
                TestAssert.AreEqual("src1", restored.SourceId);
                TestAssert.AreEqual("OneNote", restored.Host);
                TestAssert.AreEqual("{page-id}", restored.DocumentPath);
                TestAssert.AreEqual("src1", restored.TargetId);
            }
        }

        private static void ValidateCountryFlagAssets()
        {
            string root = FindRepositoryRoot();
            string flags = Path.Combine(root, "oPenEfficiency", "Resources", "Flags");
            var files = Directory.GetFiles(flags, "*.svg", SearchOption.TopDirectoryOnly);
            TestAssert.AreEqual(249, files.Length);
            foreach (var entry in CountryFlagCatalog.Entries)
                TestAssert.IsTrue(File.Exists(Path.Combine(flags, entry.AssetId + ".svg")), "Missing flag asset: " + entry.AssetId);
            foreach (string file in files)
            {
                string svg = File.ReadAllText(file);
                XDocument.Parse(svg);
                TestAssert.IsFalse(Regex.IsMatch(svg, @"(?i)<script|javascript:|<foreignObject|@import|(?:href|src)\s*=\s*['""]https?://"),
                    "Active or external content in " + Path.GetFileName(file));
            }
        }

        /// <summary>
        /// The bundled illustration packs must stay bundling-safe (no attribution obligation) and every
        /// figure must keep the named parts the picker promises: the catalog is derived from the SVG
        /// files, so a broken download or a stripped part id fails the gate instead of the user.
        /// </summary>
        private static void ValidateIllustrationAssets()
        {
            string root = FindRepositoryRoot();
            string illustrationRoot = Path.Combine(root, "oPenEfficiency", "Resources", "Illustrations");
            string manifestPath = Path.Combine(illustrationRoot, "illustrations-manifest.json");
            TestAssert.IsTrue(File.Exists(manifestPath), "Missing illustration manifest");

            var manifest = IllustrationCatalog.LoadManifest(File.ReadAllText(manifestPath));
            TestAssert.IsTrue(manifest.Collections.Count >= 5, "Illustration manifest must keep every bundled collection");

            foreach (var collection in manifest.Collections)
            {
                TestAssert.IsTrue(collection.AllowsBundling,
                    "Bundled illustration collection must allow bundling under its license: " + collection.Id);
                TestAssert.IsFalse(string.IsNullOrWhiteSpace(collection.License), "Missing license for " + collection.Id);
                TestAssert.IsFalse(string.IsNullOrWhiteSpace(collection.LicenseUrl), "Missing license url for " + collection.Id);
                TestAssert.IsFalse(string.IsNullOrWhiteSpace(collection.SourceUrl), "Missing source url for " + collection.Id);
            }

            var result = IllustrationCatalog.Build(
                manifest,
                folder => Directory.GetFiles(Path.Combine(illustrationRoot, folder), "*.svg", SearchOption.TopDirectoryOnly)
                    .Select(path => folder + "/" + Path.GetFileName(path)).ToList(),
                relative =>
                {
                    string full = Path.Combine(illustrationRoot, relative.Replace('/', Path.DirectorySeparatorChar));
                    return File.Exists(full) ? File.ReadAllText(full) : null;
                });

            TestAssert.AreEqual(0, result.MissingFiles.Count, "Missing illustration assets: " + string.Join(", ", result.MissingFiles));
            TestAssert.AreEqual(0, result.UnsafeFiles.Count, "Unsafe illustration assets: " + string.Join(", ", result.UnsafeFiles));
            TestAssert.AreEqual(421, result.Items.Count, "Expected 93 Open Peeps figures, 3 Open Doodles figures and 325 Kitbitz props");

            var kitbitz = result.Items.Where(item => item.CollectionId.StartsWith("kitbitz-", StringComparison.Ordinal)).ToList();
            TestAssert.AreEqual(325, kitbitz.Count);
            TestAssert.IsTrue(kitbitz.All(item => !item.Decomposable),
                "Kitbitz ships single props without named part groups and must not claim decomposition");
            TestAssert.IsTrue(kitbitz.All(item => item.Palette.Count >= 1), "Kitbitz prop without a palette");

            var peeps = result.Items.Where(item => item.CollectionId == "open-peeps").ToList();
            TestAssert.AreEqual(93, peeps.Count);
            foreach (var peep in peeps)
            {
                TestAssert.IsTrue(peep.Decomposable, "Peep without decomposable parts: " + peep.Id);
                TestAssert.IsTrue(peep.Parts.Any(part => part.StartsWith("head/", StringComparison.Ordinal)),
                    "Peep without a named head part: " + peep.Id);
                TestAssert.IsTrue(peep.Palette.Count >= 1, "Peep without a palette: " + peep.Id);
            }

            TestAssert.IsTrue(result.Items.Where(item => item.CollectionId == "open-doodles").All(item => !item.Decomposable),
                "Open Doodles is a recolour-only collection and must not claim part decomposition");

            foreach (string file in Directory.GetFiles(illustrationRoot, "*.svg", SearchOption.AllDirectories))
            {
                string svg = File.ReadAllText(file);
                XDocument.Parse(svg);
                TestAssert.IsFalse(Regex.IsMatch(svg, @"(?i)<script|javascript:|<foreignObject|@import|<image|(?:href|src)\s*=\s*['""]https?://"),
                    "Active or external content in " + Path.GetFileName(file));
            }
        }

        private static void ValidatePerformanceModeIdentitySelection()
        {
            TestAssert.IsTrue(ExternalAddInIdentity.IsOpenEfficiency("oPenEfficiency", null));
            TestAssert.IsTrue(ExternalAddInIdentity.IsOpenEfficiency("Vendor.AddIn", "open Efficiency PowerPoint tools"));
            TestAssert.IsFalse(ExternalAddInIdentity.IsOpenEfficiency("Vendor.AddIn", "External helper"));
            CollectionAssert.AreEqual(
                new[] { "Alpha.AddIn", "Beta.AddIn" },
                ExternalAddInIdentity.Normalize(new[] { " Beta.AddIn ", "Alpha.AddIn", "alpha.addin", "oPenEfficiency" }).ToArray());
        }

        private static void ValidateShapeFillPaletteMath()
        {
            TestAssert.AreEqual(0x336699, ShapeFillPalette.ApplyTint(0x336699, 0f));
            TestAssert.AreEqual(0xFFFFFF, ShapeFillPalette.ApplyTint(0x336699, 1f));
            TestAssert.AreEqual(0x000000, ShapeFillPalette.ApplyTint(0x336699, -1f));
            TestAssert.IsTrue(ShapeFillPalette.TryParseHex("#12AbEF", out int rgb));
            TestAssert.AreEqual(0x12ABEF, rgb);
            TestAssert.IsFalse(ShapeFillPalette.TryParseHex("not-a-color", out rgb));
            TestAssert.AreEqual(10, ShapeFillPalette.StandardRgb.Length);
        }

        private static void ValidateQuickContextMenuConfiguration()
        {
            var known = new HashSet<string>(new[] { "A", "B", "C" }, StringComparer.Ordinal);
            var normalized = QuickContextMenuSettings.Normalize(
                new[] { "B", "", "B", "missing", "A", "C" },
                known);
            CollectionAssert.AreEqual(new[] { "B", "A", "C" }, normalized);

            var many = Enumerable.Range(0, 20).Select(index => "F" + index).ToArray();
            var allKnown = new HashSet<string>(many, StringComparer.Ordinal);
            TestAssert.AreEqual(
                QuickContextMenuSettings.MaximumItems,
                QuickContextMenuSettings.Normalize(many, allKnown).Count,
                "Quick Context Menu configuration must remain bounded.");
        }

        private static void ValidateCopyPasteXyPositionOnly()
        {
            string root = FindRepositoryRoot();
            string copySource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Features", "Utilities", "CopyCoordinatesFeature.cs"));
            string pasteSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Features", "Utilities", "PasteCoordinatesFeature.cs"));
            string policySource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency.Core", "UI", "FeatureCatalogPolicy.cs"));
            string copyMenuSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "MainSidebar.ContextMenus.Layout.cs"));
            string pasteMenuSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "MainSidebar.ContextMenus.WorkflowAndVisual.cs"));
            string positionPainterSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Features", "Visuals", "PositionPainterFeature.cs"));

            string copyExecute = ExtractMethodBody(
                copySource,
                "public static bool Execute(PowerPointManager manager)",
                "CopyCoordinatesFeature.Execute");
            string clearCopiedGeometry = ExtractMethodBody(
                copySource,
                "public static bool ClearCopiedGeometry()",
                "CopyCoordinatesFeature.ClearCopiedGeometry");
            string pasteExecute = ExtractMethodBody(
                pasteSource,
                "public static bool Execute(PowerPointManager manager)",
                "PasteCoordinatesFeature.Execute");

            int copySerialize = copyExecute.IndexOf("CoordinateClipboardPayload.TrySerializeCurrent", StringComparison.Ordinal);
            int copyClipboardSet = copyExecute.IndexOf("System.Windows.Forms.Clipboard.SetText", StringComparison.Ordinal);
            Assert(copySerialize >= 0 && copyClipboardSet > copySerialize,
                "CopyCoordinatesFeature.Execute must serialize a current owned payload before writing clipboard text.");
            Assert(!Regex.IsMatch(copyExecute, @"shapes\s*\[\s*1\s*\]\s*\.\s*(Width|Height)"),
                "Copy XY must not read Width or Height.");
            int clearOwnershipCheck = clearCopiedGeometry.IndexOf("CoordinateClipboardPayload.IsOwnedCurrentPayload", StringComparison.Ordinal);
            int clipboardClear = clearCopiedGeometry.IndexOf("System.Windows.Forms.Clipboard.Clear", StringComparison.Ordinal);
            Assert(clearOwnershipCheck >= 0 && clipboardClear > clearOwnershipCheck,
                "ClearCopiedGeometry must prove current-payload ownership before clearing the clipboard.");
            Assert(!clearCopiedGeometry.Contains("CoordinateClipboardPayload.TryParseForPaste"),
                "Clear Copied XY Coordinates must not use the paste-compatible parser as an ownership check.");
            int pasteParser = pasteExecute.IndexOf("CoordinateClipboardPayload.TryParseForPaste", StringComparison.Ordinal);
            MatchCollection coordinateAssignments = Regex.Matches(
                pasteExecute,
                @"shapes\s*\[\s*1\s*\]\s*\.\s*(Left|Top)\s*=",
                SourceOptions);
            Assert(pasteParser >= 0 && coordinateAssignments.Count == 2 &&
                   coordinateAssignments.Cast<Match>().All(match => pasteParser < match.Index),
                "PasteCoordinatesFeature.Execute must parse the payload before assigning Left or Top.");
            Assert(!Regex.IsMatch(pasteExecute, @"shapes\s*\[\s*1\s*\]\s*\.\s*(Width|Height)\s*="),
                "Paste XY must never assign Width or Height.");
            Assert(!copySource.Contains("CopyMode") && !pasteSource.Contains("PasteMode"),
                "Copy/Paste XY must not retain size or both execution modes.");

            Assert(Regex.IsMatch(policySource,
                    @"\[""BtnCopyXY""\]\s*=\s*new\[\]\s*\{\s*""Clear Copied XY Coordinates""\s*\}"),
                "Copy XY catalog variants must contain only the recognized-XY clear action.");
            Assert(Regex.IsMatch(policySource,
                    @"\[""BtnPasteXY""\]\s*=\s*new\[\]\s*\{\s*""Clear Copied XY Coordinates""\s*\}"),
                "Paste XY catalog variants must contain only the recognized-XY clear action.");
            Assert(!copyMenuSource.Contains("CopyCoordinatesFeature.CopyMode"),
                "Copy XY context menu must not expose size or both execution modes.");
            Assert(!pasteMenuSource.Contains("PasteCoordinatesFeature.PasteMode"),
                "Paste XY context menu must not expose size or both execution modes.");
            Assert(!copyMenuSource.Contains("Copy Size (Width/Height)") &&
                   !copyMenuSource.Contains("Copy Position and Size"),
                "Copy XY context menu must not offer size or both variants.");
            Assert(!pasteMenuSource.Contains("Paste Size (Width/Height)") &&
                   !pasteMenuSource.Contains("Paste Both (Coordinates & Size)"),
                "Paste XY context menu must not offer size or both variants.");
            Assert(positionPainterSource.Contains("applyWidth") && positionPainterSource.Contains("applyHeight"),
                "PositionPainter must remain the separate workflow that transfers dimensions.");
        }

        private static void ValidateFeatureIconGlyphHitTesting()
        {
            string root = FindRepositoryRoot();
            string xaml = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Controls", "FeatureIcon.xaml"));
            Assert(xaml.Contains("IconLayerHost") && xaml.Contains("IsHitTestVisible=\"False\""),
                "FeatureIcon glyph layer must stay IsHitTestVisible=False: state changes rebuild the " +
                "paths during mouse-down dispatch, and a hit element destroyed mid-event silently swallows the click.");
        }

        private static void ValidateMouseHookCallbacksAreComFree()
        {
            string root = FindRepositoryRoot();
            foreach (string managerPath in new[]
            {
                Path.Combine("oPenEfficiency", "Services", "SnapToGridManager.cs"),
                Path.Combine("oPenEfficiency", "Services", "SnapToObjectsManager.cs")
            })
            {
                string source = File.ReadAllText(Path.Combine(root, managerPath));
                foreach (string handler in new[] { "OnGlobalMouseDown", "OnGlobalMouseUp", "OnMouseMoveFilter" })
                {
                    int asVoid = source.IndexOf("void " + handler + "(", StringComparison.Ordinal);
                    int asBool = source.IndexOf("bool " + handler + "(", StringComparison.Ordinal);
                    int start = asVoid < 0 ? asBool : (asBool < 0 ? asVoid : Math.Min(asVoid, asBool));
                    Assert(start >= 0, managerPath + " no longer exposes " + handler + "; update the PQ-015 contract.");
                    int bodyEnd = source.IndexOf("\r\n        }\r\n", start, StringComparison.Ordinal);
                    if (bodyEnd < 0) bodyEnd = source.IndexOf("\n        }\n", start, StringComparison.Ordinal);
                    Assert(bodyEnd > start, managerPath + "." + handler + " body end could not be isolated.");
                    string body = source.Substring(start, bodyEnd - start);
                    Assert(!body.Contains("_powerPointManager."),
                        managerPath + "." + handler + " still calls PowerPoint COM synchronously inside the mouse-hook callback (PQ-015).");
                }
                string downStart = source.Substring(source.IndexOf("private void OnGlobalMouseDown(", StringComparison.Ordinal));
                int downEnd = downStart.IndexOf("private void ArmDragFromSelection", StringComparison.Ordinal);
                Assert(downEnd > 0, managerPath + " does not defer mouse-down COM to the dispatcher.");
                string downBody = downStart.Substring(0, downEnd);
                Assert(downBody.Contains("BeginInvoke"),
                    managerPath + " mouse-down path lost its dispatcher deferral.");
            }

            string coordinator = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Services", "GanttInteractionCoordinator.cs"));
            int coordDown = coordinator.IndexOf("private void OnLeftButtonDown(", StringComparison.Ordinal);
            int coordDownEnd = coordinator.IndexOf("private void OnLeftButtonUp", coordDown, StringComparison.Ordinal);
            Assert(coordDown >= 0 && coordDownEnd > coordDown, "Could not inspect the Gantt coordinator mouse-down path.");
            string coordDownBody = coordinator.Substring(coordDown, coordDownEnd - coordDown);
            Assert(coordDownBody.Contains("AnyGanttWindowOpen"),
                "Gantt mouse-down probing is not gated on an open Gantt window.");
            Assert(coordDownBody.IndexOf("DeferToUi", StringComparison.Ordinal) <
                   coordDownBody.IndexOf("_interactionManager.OnMouseDown", StringComparison.Ordinal),
                "Gantt mouse-down COM still runs synchronously in the hook callback.");
            Assert(coordDownBody.Contains("IsLeftMouseButtonDown"),
                "Deferred Gantt mouse-down work lacks the button-still-held guard.");
            int coordRight = coordinator.IndexOf("private void OnHookRightButtonUp", StringComparison.Ordinal);
            Assert(coordRight >= 0 && coordinator.Substring(coordRight).Contains("DeferToUi"),
                "Gantt right-button COM is not deferred out of the hook callback.");

            string hook = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Utils", "Win32MouseHook.cs"));
            Assert(hook.Contains("internal static bool IsLeftMouseButtonDown()") &&
                   hook.Contains("\"Win32MouseHook.WM_LBUTTONUP\""),
                "Win32MouseHook lost its button-state helper or its WM_LBUTTONUP exception guard.");

            // PQ-016: exactly one thread hook per process, owned by ThisAddIn.
            foreach (string managerPath in new[]
            {
                Path.Combine("oPenEfficiency", "Services", "SnapToGridManager.cs"),
                Path.Combine("oPenEfficiency", "Services", "SnapToObjectsManager.cs"),
                Path.Combine("oPenEfficiency", "Services", "GanttInteractionCoordinator.cs")
            })
            {
                string source = File.ReadAllText(Path.Combine(root, managerPath));
                Assert(!System.Text.RegularExpressions.Regex.IsMatch(
                        source, @"new\s+[\w.]*Win32MouseHook\s*\("),
                    managerPath + " installs its own WH_MOUSE hook; managers must subscribe to the " +
                    "single ThisAddIn-owned shared hook instead.");
            }
            string addIn = File.ReadAllText(Path.Combine(root, "oPenEfficiency", "ThisAddIn.cs"));
            int sharedCreations = System.Text.RegularExpressions.Regex.Matches(
                addIn, @"new oPenEfficiency\.Utils\.Win32MouseHook\(\)").Count;
            Assert(sharedCreations == 1,
                "ThisAddIn must create exactly one shared Win32MouseHook (found " + sharedCreations + ").");
            Assert(addIn.Contains("_sharedMouseHook.Dispose()"),
                "ThisAddIn must dispose the shared mouse hook after its subscribers during shutdown.");
        }

        private static void ValidateSelectionDebounceDoesNotRetainTransientComObjects()
        {
            string root = FindRepositoryRoot();
            string coordinator = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Services", "SelectionUpdateCoordinator.cs"));
            string addIn = File.ReadAllText(Path.Combine(root, "oPenEfficiency", "ThisAddIn.cs"));
            string sidebar = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "MainSidebar.xaml.cs"));

            Assert(!coordinator.Contains("Microsoft.Office.Interop") &&
                   !coordinator.Contains("PowerPoint.Selection") &&
                   coordinator.Contains("Action _processSelection") &&
                   coordinator.Contains("DispatcherPriority.ApplicationIdle"),
                "Selection debounce must retain only a refresh signal, never PowerPoint's transient Selection RCW.");
            Assert(addIn.Contains("_selectionUpdates?.Schedule();") &&
                   addIn.Contains("ProcessCurrentSelectionChange") &&
                   addIn.Contains("activeWindow?.Selection"),
                "Selection processing must reacquire the current Selection after the debounce interval.");
            Assert(!sidebar.Contains("_currentSelection"),
                "Sidebar must not retain a Selection RCW beyond the current refresh pass.");
        }

        private static void ValidateRegisteredSidebarDialogFailurePaths()
        {
            string root = FindRepositoryRoot();
            var boundaries = new[]
            {
                new { Path = Path.Combine("oPenEfficiency", "Features", "Utilities", "AgendaWizardFeature.cs"), Context = "AgendaWizardFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Alignment", "AlignToTableCellFeature.cs"), Context = "AlignToTableCellFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Text", "BulkFormatReplacerFeature.cs"), Context = "BulkFormatReplacerFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Utilities", "DocumentAutomationFeature.cs"), Context = "DocumentAutomationFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Utilities", "ExcelLinkManagerFeature.cs"), Context = "ExcelLinkManagerFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Tables", "TableFormatPainterFeature.cs"), Context = "TableFormatPainterFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Tables", "TableBrandingFeature.cs"), Context = "TableBrandingFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Tables", "TableHeatmapFeature.cs"), Context = "TableHeatmapFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Text", "FormatNumbersFeature.cs"), Context = "FormatNumbersFeature.ExecuteDialog" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Utilities", "SaveAgendaLayoutFeature.cs"), Context = "SaveAgendaLayoutFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Utilities", "SlidePasteFeature.cs"), Context = "SlidePasteFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Visuals", "MapWizardFeature.cs"), Context = "MapWizardFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Visuals", "PositionPainterFeature.cs"), Context = "PositionPainterFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Visuals", "TransparentColorFeature.cs"), Context = "TransparentColorFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Alignment", "DiagonalAlignFeature.cs"), Context = "DiagonalAlignFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Utilities", "SlideSizeFeature.cs"), Context = "SlideSizeFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Utilities", "StorylineFeature.cs"), Context = "StorylineFeature.OpenFiltered" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Utilities", "TagInspectorFeature.cs"), Context = "TagInspectorFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Visuals", "ThemeColorPickerFeature.cs"), Context = "ThemeColorPickerFeature.Execute" },
                new { Path = Path.Combine("oPenEfficiency", "Features", "Visuals", "IconBrowserFeature.cs"), Context = "IconBrowserFeature.Execute" }
            };

            foreach (var boundary in boundaries)
            {
                string source = File.ReadAllText(Path.Combine(root, boundary.Path));
                string loggedBoundary = "ExceptionLogger\\.Log\\(ex,\\s*\"" +
                    Regex.Escape(boundary.Context) + "\"\\);";
                Assert(!Regex.IsMatch(
                        source,
                        loggedBoundary + @"\s*return false;",
                        SourceOptions),
                    boundary.Path + " still returns silent false immediately after logging a launch failure.");
                Assert(Regex.IsMatch(
                        source,
                        loggedBoundary + @"[\s\S]*?throw;",
                        SourceOptions),
                    boundary.Path + " still converts a window construction/show exception into silent false.");
            }
        }

        private static void ValidateSidebarNavigationFailurePath()
        {
            string root = FindRepositoryRoot();
            string navigation = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "SidebarNavigationService.cs"));
            string addIn = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "ThisAddIn.cs"));

            Assert(!navigation.Contains("Topmost = true"),
                "Sidebar navigation still uses Topmost instead of the PowerPoint owner.");
            Assert(Regex.IsMatch(
                    navigation,
                    @"catch \(Exception ex\)[\s\S]*?SidebarNavigationService\.OpenSpecialCharacters[\s\S]*?throw;",
                    RegexOptions.CultureInvariant),
                "Special Characters window failures are logged but swallowed as command success.");
            Assert(Regex.IsMatch(
                    navigation,
                    @"OpenTemplateManager\(\)[\s\S]*?if \(manager == null\)\s*throw new InvalidOperationException",
                    RegexOptions.CultureInvariant),
                "Template Manager silently returns when its PowerPoint manager lookup fails.");

            int explorerStart = addIn.IndexOf("public void OpenFeatureExplorer(", StringComparison.Ordinal);
            int explorerEnd = addIn.IndexOf("public void OpenIconBrowser(", explorerStart, StringComparison.Ordinal);
            Assert(explorerStart >= 0 && explorerEnd > explorerStart,
                "Could not inspect the Feature Explorer host boundary.");
            string explorerBody = addIn.Substring(explorerStart, explorerEnd - explorerStart);
            Assert(Regex.IsMatch(
                    explorerBody,
                    @"catch \(Exception ex\)[\s\S]*?ExceptionLogger\.Log\([^;]+;[\s\S]*?throw;",
                    RegexOptions.CultureInvariant),
                "Feature Explorer construction/show failures are logged but swallowed as command success.");

            int iconBrowserStart = addIn.IndexOf("public void OpenIconBrowser(", StringComparison.Ordinal);
            int iconBrowserEnd = addIn.IndexOf("public void ToggleSidebar(", iconBrowserStart, StringComparison.Ordinal);
            Assert(iconBrowserStart >= 0 && iconBrowserEnd > iconBrowserStart,
                "Could not inspect the Icon Browser host boundary.");
            string iconBrowserBody = addIn.Substring(iconBrowserStart, iconBrowserEnd - iconBrowserStart);
            Assert(Regex.IsMatch(
                    iconBrowserBody,
                    @"catch \(Exception ex\)[\s\S]*?ExceptionLogger\.Log\([^;]+;[\s\S]*?throw;",
                    RegexOptions.CultureInvariant),
                "Icon Browser construction/show failures are logged but swallowed as command success.");

            int assetStart = addIn.IndexOf("public void ToggleAssetManager()", StringComparison.Ordinal);
            int assetEnd = addIn.IndexOf("protected override Microsoft.Office.Core.IRibbonExtensibility", assetStart, StringComparison.Ordinal);
            Assert(assetStart >= 0 && assetEnd > assetStart,
                "Could not inspect the Asset Manager host boundary.");
            string assetBody = addIn.Substring(assetStart, assetEnd - assetStart);
            Assert(assetBody.Contains("ExceptionLogger.Log(ex, \"ThisAddIn.ToggleAssetManager\")") &&
                   assetBody.Contains("throw;") && !assetBody.Contains("ex.Message"),
                "Asset Manager failures are not safely logged and propagated to Sidebar completion.");
        }

        private static void ValidateSidebarManagerFailurePath()
        {
            string root = FindRepositoryRoot();
            string sidebar = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "MainSidebar.xaml.cs"));
            string commands = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "MainSidebar.Commands.cs"));

            int primaryStart = sidebar.IndexOf("private void ExecuteFeature(", StringComparison.Ordinal);
            int primaryEnd = sidebar.IndexOf("private void OpenApplyTextEditor", primaryStart, StringComparison.Ordinal);
            Assert(primaryStart >= 0 && primaryEnd > primaryStart,
                "Could not inspect the primary Sidebar click path.");
            string primaryBody = sidebar.Substring(primaryStart, primaryEnd - primaryStart);
            int primaryUiCreation = primaryBody.IndexOf("new SidebarCommandUiAdapter", StringComparison.Ordinal);
            int primaryManagerCheck = primaryBody.IndexOf("if (manager == null)", StringComparison.Ordinal);
            Assert(primaryUiCreation >= 0 && primaryUiCreation < primaryManagerCheck,
                "The primary Sidebar click path creates no UI adapter before the PowerPoint-manager failure path.");
            Assert(Regex.IsMatch(
                    primaryBody,
                    @"if \(manager == null\)[\s\S]*?CompleteRegisteredCommand\([^;]+;[\s\S]*?return",
                    RegexOptions.CultureInvariant),
                "The primary Sidebar click path silently returns when PowerPoint is unavailable.");
            Assert(primaryBody.Contains("for '{id}' from {source}"),
                "Primary manager-unavailable diagnostics omit command ID or invocation source.");

            int start = commands.IndexOf("private CommandDispatchResult ExecuteRegisteredCommand", StringComparison.Ordinal);
            int end = commands.IndexOf("private static ToggleIntent GetToggleIntent", start, StringComparison.Ordinal);
            Assert(start >= 0 && end > start, "Could not inspect Sidebar command completion.");
            string body = commands.Substring(start, end - start);

            int uiCreation = body.IndexOf("new SidebarCommandUiAdapter", StringComparison.Ordinal);
            int managerCheck = body.IndexOf("if (manager == null)", StringComparison.Ordinal);
            Assert(uiCreation >= 0 && uiCreation < managerCheck,
                "Sidebar creates no UI adapter before the PowerPoint-manager failure path.");
            Assert(Regex.IsMatch(
                    body,
                    @"if \(manager == null\)[\s\S]*?CompleteRegisteredCommand\([^;]+;[\s\S]*?return",
                    RegexOptions.CultureInvariant),
                "PowerPoint-manager failure is returned without user-visible completion.");
            Assert(body.Contains("for '{key.Id}' from {source}"),
                "Parameterized manager-unavailable diagnostics omit command ID or invocation source.");
        }

        private static void ValidateSidebarModelessWindowLifecycle()
        {
            string root = FindRepositoryRoot();
            string sidebar = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "MainSidebar.xaml.cs"));
            string commands = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "MainSidebar.Commands.cs"));
            int start = commands.IndexOf("public bool SetModelessWindowState(", StringComparison.Ordinal);
            int end = commands.IndexOf(
                "public CommandDispatchResult ExecuteHostAction(",
                start,
                StringComparison.Ordinal);
            Assert(start >= 0 && end > start,
                "Could not inspect the modeless Sidebar window lifecycle.");

            string body = commands.Substring(start, end - start);
            Assert(body.Contains("WindowInteropHelper(window).Owner"),
                "Modeless Sidebar windows are not owned by the active PowerPoint HWND.");
            Assert(body.IndexOf("window.Show();", StringComparison.Ordinal) <
                   body.IndexOf("_floatingWindows[commandId] = window;", StringComparison.Ordinal),
                "A modeless window is tracked before Show succeeds, leaving stale entries on failure.");
            Assert(body.Contains("!window.IsLoaded || !window.IsVisible"),
                "A modeless window that closes during Show can still be retained as a stale entry.");
            Assert(body.Contains("!existing.IsLoaded") && body.Contains("Remove(commandId)"),
                "Stale unloaded modeless-window entries are not removed before retry.");
            Assert(body.Contains("existing.WindowState == WindowState.Minimized") &&
                   body.Contains("existing.WindowState = WindowState.Normal"),
                "An existing minimized Sidebar tool is not restored before activation.");
            Assert(commands.Contains("throw new InvalidOperationException(\"Sidebar host is unavailable"),
                "Host-backed Sidebar commands can still report success when no host is available.");

            int ganttStart = sidebar.IndexOf("private void ShowOrCenterGanttWindow(", StringComparison.Ordinal);
            int ganttEnd = sidebar.IndexOf("private void CenterWindowOverSlide", ganttStart, StringComparison.Ordinal);
            Assert(ganttStart >= 0 && ganttEnd > ganttStart,
                "Could not inspect the Gantt modeless-window lifecycle.");
            string ganttBody = sidebar.Substring(ganttStart, ganttEnd - ganttStart);
            Assert(ganttBody.IndexOf("window.Show();", StringComparison.Ordinal) <
                   ganttBody.IndexOf("_floatingWindows[key] = window;", StringComparison.Ordinal),
                "Gantt tracks its window before Show succeeds.");
            Assert(ganttBody.Contains("WindowState.Minimized") && ganttBody.Contains("WindowState.Normal") &&
                   ganttBody.Contains(".Focus();"),
                "Gantt does not restore and focus an existing minimized window.");
            Assert(ganttBody.Contains("catch") && ganttBody.Contains("_floatingWindows.Remove(key)") &&
                   ganttBody.Contains("throw;"),
                "Gantt does not clean up and propagate a failed Show.");

            int genericStart = sidebar.IndexOf("private void ToggleFloatingWindow<T>(", StringComparison.Ordinal);
            int genericEnd = sidebar.IndexOf("internal bool ExerciseFloatingWindowLifecycleForAutomation", genericStart, StringComparison.Ordinal);
            Assert(genericStart >= 0 && genericEnd > genericStart,
                "Could not inspect the generic modeless-window lifecycle.");
            string genericBody = sidebar.Substring(genericStart, genericEnd - genericStart);
            Assert(genericBody.IndexOf("window.Show();", StringComparison.Ordinal) <
                   genericBody.IndexOf("_floatingWindows[key] = window;", StringComparison.Ordinal),
                "Generic floating-window commands track before Show succeeds.");
            Assert(genericBody.Contains("WindowState.Minimized") && genericBody.Contains("WindowState.Normal") &&
                   genericBody.Contains(".Focus();") && genericBody.Contains("throw;"),
                "Generic floating-window commands do not restore/activate or propagate Show failures safely.");

            string arrangeGrid = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Features", "Alignment", "ArrangeGridFeature.cs"));
            Assert(arrangeGrid.Contains("_window.WindowState == System.Windows.WindowState.Minimized") &&
                   arrangeGrid.Contains("_window.Focus();") && arrangeGrid.Contains("throw;"),
                "Arrange Grid does not restore/focus existing UI or propagate launch failures safely.");

            string syncObjects = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Features", "Visuals", "SyncObjectsFeature.cs"));
            Assert(syncObjects.Contains("private static SyncObjectsWindow _window") &&
                   syncObjects.Contains("WindowInteropHelper(_window).Owner") &&
                   syncObjects.Contains("_window.WindowState == System.Windows.WindowState.Minimized") &&
                   syncObjects.Contains("_window.Focus();") && syncObjects.Contains("throw;"),
                "Sync Objects lacks owned single-instance modeless-window lifecycle handling.");
        }

        private static void ValidateSidebarThemeTrackingContract()
        {
            string root = FindRepositoryRoot();
            string sidebar = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "MainSidebar.xaml.cs"));
            string settings = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "SettingsViewModel.cs"));
            string settingsXaml = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Dialogs", "SettingsWindow.xaml"));
            string config = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency.Core", "Models", "SidebarConfig.cs"));

            Assert(sidebar.Contains("SidebarBackgroundModes.Normalize(config)"),
                "Sidebar does not migrate background tracking to an explicit state.");
            Assert(sidebar.Contains(
                    "SetResourceReference(Control.BackgroundProperty, \"SidebarBackgroundBrush\")"),
                "Sidebar theme tracking does not use the dedicated Sidebar background token.");
            Assert(Regex.IsMatch(
                    sidebar,
                    @"ThemeManager_ThemeChanged\([^)]*\)\s*\{[^}]*RebuildUI\(_activeSearchTerm\);",
                    SourceOptions),
                "An open Sidebar does not preserve its active search when the application theme changes.");
            int rebuildStart = sidebar.IndexOf("public void RebuildUI(", StringComparison.Ordinal);
            int rebuildEnd = sidebar.IndexOf("private UIElement GetWidgetForId", rebuildStart, StringComparison.Ordinal);
            Assert(rebuildStart >= 0 && rebuildEnd > rebuildStart,
                "Could not inspect the Sidebar rebuild path.");
            string rebuildBody = sidebar.Substring(rebuildStart, rebuildEnd - rebuildStart);
            Assert(!rebuildBody.Contains("config = new SidebarConfig()"),
                "Sidebar search mode still discards persisted custom appearance settings.");
            Assert(rebuildBody.IndexOf("JsonService.Load<SidebarConfig>", StringComparison.Ordinal) <
                   rebuildBody.IndexOf("if (isSearchMode)", StringComparison.Ordinal),
                "Sidebar search mode does not load persisted appearance before creating search results.");
            Assert(settings.Contains("UseThemeBackground"),
                "Settings does not expose explicit theme-tracking versus custom Sidebar background state.");
            Assert(settings.Contains("exportData.BackgroundMode = null") &&
                   settings.Contains("exportData.BackgroundMode = Config.BackgroundMode"),
                "Portable settings export does not preserve the explicit Sidebar background mode.");
            Assert(settings.Contains("Config.BackgroundMode = imported.BackgroundMode"),
                "Portable settings import does not restore the explicit Sidebar background mode.");
            Assert(settingsXaml.Contains("Follow app theme"),
                "Settings does not present the Sidebar background mode to the user.");
            Assert(!config.Contains("IsDarkMode"),
                "SidebarConfig still duplicates the authoritative AppTheme state.");
        }

        private static void ValidateCleanerSanitizationActions()
        {
            string root = FindRepositoryRoot();
            string feature = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Features", "Utilities", "CleanerFeature.cs"));
            string menu = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "MainSidebar.ContextMenus.WorkflowAndVisual.cs"));

            foreach (string action in new[]
            {
                "RemoveAlternativeText", "UnlockShapes", "RemoveInk", "BreakDataLinks",
                "RemoveInvisibleOrOffSlide", "RemoveDocumentInformation", "RemoveUnusedLayouts",
                "RemoveNonOpeCustomXml", "MakeInvisibleShapesVisible", "SanitizeNumbers"
            })
                Assert(feature.Contains(action), "Cleaner is missing action: " + action);
            Assert(feature.Contains("ShapeMetadataService.HasOpeTags(shape)") &&
                   feature.Contains("ShapeMetadataService.HasThirdPartyTags(shape)"),
                "Destructive Cleaner shape actions no longer preserve owned/protected smart objects.");
            Assert(menu.Contains("Remove Non-oPE Custom XML") && menu.Contains("Sanitize Visible Numbers"),
                "New Cleaner actions are not discoverable from the Cleaner context menu.");
        }

        private static void ValidateConnectedTextLabels()
        {
            string root = FindRepositoryRoot();
            string feature = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Features", "Utilities", "ConnectedTextLabelFeature.cs"));
            string menu = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "MainSidebar.RemainingContextMenus.cs"));
            string readme = File.ReadAllText(Path.Combine(root, "README.md"));
            string client = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Services", "OneNote", "OneNoteDesktopClient.cs"));
            string pageUpdates = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Services", "OneNote", "OneNotePageUpdateService.cs"));
            string bridgeClient = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Services", "OneNote", "OneNoteBridgeClient.cs"));
            string mutations = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency.Core", "OneNote", "OneNotePageMutations.cs"));
            string store = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Services", "OneNote", "ConnectedOneNoteTargetStore.cs"));

            Assert(feature.Contains("ContentControls.Add") && feature.Contains("OneNotePageUpdateService.TryAppendConnectedText"),
                "Insert Linked Text Label no longer creates both Word and OneNote targets.");
            Assert(feature.Contains("oPE_ConnectedTextSourceId") &&
                   store.Contains("connected_text_labels.json") &&
                   feature.Contains("RefreshAll"),
                "Insert Linked Text Label no longer persists source/target identities for refresh.");
            // FEAT-007 ON-S0 safety boundary: active page resolution and no-force updates.
            Assert(client.Contains("CurrentPageId") && !client.Contains("GetHierarchy"),
                "OneNote targets must resolve the active page through CurrentPageId, never GetHierarchy.");
            Assert(client.Contains("Process.GetProcessesByName(\"ONENOTE\")") &&
                   client.Contains("Activator.CreateInstance(applicationType)"),
                "OneNote connection must require a visible desktop window and support builds that are absent from the ROT.");
            Assert(client.Contains("expectedLastModified"),
                "OneNote page updates must pass an expected-last-modified token.");
            Assert(client.Contains("expectedLastModified == DateTime.MinValue"),
                "OneNote page updates must fail closed when no concurrency token is available.");
            Assert(pageUpdates.Contains("OneNotePageMutations.TryCreateAppend") &&
                   mutations.Contains("changedOutlines.Select(outline => new XElement(outline))"),
                "OneNote updates must submit only changed openEfficiency page-level objects.");
            Assert(bridgeClient.Contains("OneNoteBridgeProtocol.GetActivePage") &&
                   feature.Contains("bridge.TryAppendConnectedText"),
                "PowerPoint must fall back to the typed in-process OneNote bridge.");
            string publicSolution = File.ReadAllText(Path.Combine(root, "oPenEfficiency.slnx"));
            Assert(!publicSolution.Contains("oPenEfficiency.OneNote/"),
                "The paused OneNote host must remain outside the public solution.");
            Assert(!client.Contains(", true);") && !feature.Contains(", true);"),
                "Unconditional force updates are not allowed in the OneNote update path.");
            Assert(menu.Contains("Send to Word") && menu.Contains("Send to OneNote") &&
                   menu.Contains("Refresh All Open Targets"),
                "Connected Text Label target and refresh actions are not discoverable.");
            Assert(readme.Contains("Insert Linked Text Label"),
                "README does not document Insert Linked Text Label.");
        }

        private static void ValidateProgressSeriesColorDetection()
        {
            string root = FindRepositoryRoot();
            string toolbar = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Toolbars", "ProgressSeriesToolbar.xaml.cs"));

            Assert(toolbar.Contains("seriesGroup.GroupItems[_settings.ActiveIndex]") &&
                   toolbar.Contains("ColorConstants.ToHex(activeShape.Fill.ForeColor.RGB)"),
                "Progress Series does not read the visible active-segment fill color.");
            Assert(!toolbar.Contains("TODO: Implement safer color detection without ungrouping") &&
                   !toolbar.Contains(".Ungroup("),
                "Progress Series color detection is still deferred or destructively ungroups the series.");
        }

        private static void ValidateGroupedChartOverlays()
        {
            string root = FindRepositoryRoot();
            string feature = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Features", "Charts", "ChartOverlayFeature.cs"));
            string dialog = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Dialogs", "ChartOverlayDialog.xaml.cs"));
            string readme = File.ReadAllText(Path.Combine(root, "README.md"));

            Assert(feature.Contains("FindSingleChart(selection[1])") &&
                   feature.Contains("shape.GroupItems[index]") &&
                   feature.Contains("chartCount == 1"),
                "Add Analytical Chart Overlays no longer resolves exactly one chart inside a selected group.");
            Assert(dialog.Contains("_manager.GetCurrentSlide()"),
                "Grouped chart overlay attempts to cast the chart's group parent to a slide.");
            Assert(readme.Contains("one group containing exactly one native chart"),
                "README does not disclose grouped-chart overlay support.");
        }

        private static void ValidateGanttRulesModeToggle()
        {
            string root = FindRepositoryRoot();
            string xaml = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "GanttChartPanel.xaml"));
            var toggle = Regex.Match(
                xaml,
                @"<CheckBox(?=[^>]*IsChecked=""\{Binding Project\.FollowGanttRules\}""" +
                @")(?=[^>]*Content=""Follow Gantt Rules"")[^>]*/>",
                SourceOptions);

            Assert(toggle.Success,
                "The Follow Gantt Rules mode toggle is missing from the Gantt panel.");
            Assert(!toggle.Value.Contains("IsEnabled="),
                "The Follow Gantt Rules toggle disables itself and cannot be re-enabled.");
        }

        private static void ValidateShortcutSafety(string root)
        {
            var empty = ShortcutPolicy.Validate(Array.Empty<string>(), riskConfirmed: true);
            Assert(!empty.IsValid, "An empty shortcut chord was accepted.");

            var normalized = ShortcutPolicy.Validate(
                new[] { "RightAlt", "RightCtrl", "M" },
                riskConfirmed: true);
            Assert(normalized.IsValid &&
                   normalized.Chord.Signature == "Ctrl|Alt|M" &&
                   normalized.Chord.DisplayText == "Ctrl + Alt + M",
                "Left/right modifier normalization is not canonical.");

            var leftNormalized = ShortcutPolicy.Validate(
                new[] { "LeftCtrl", "LeftAlt", "M" },
                riskConfirmed: true);
            Assert(leftNormalized.Chord.Signature == normalized.Chord.Signature,
                "Left and right modifier variants resolve differently.");

            var duplicate = ShortcutPolicy.Validate(
                new[] { "LeftCtrl", "RightCtrl", "M" },
                riskConfirmed: true);
            Assert(!duplicate.IsValid &&
                   duplicate.Error.IndexOf("more than once", StringComparison.OrdinalIgnoreCase) >= 0,
                "Duplicate normalized keys were accepted.");

            var unsupported = ShortcutPolicy.Validate(
                new[] { "ImeProcessed", "A" },
                riskConfirmed: true);
            Assert(!unsupported.IsValid &&
                   unsupported.Error.IndexOf("reliably", StringComparison.OrdinalIgnoreCase) >= 0,
                "An unsupported runtime key was accepted.");

            var reservedWindows = ShortcutPolicy.Validate(
                new[] { "LWin", "D" },
                riskConfirmed: true);
            Assert(!reservedWindows.IsValid &&
                   reservedWindows.Error.IndexOf("Windows-key", StringComparison.OrdinalIgnoreCase) >= 0,
                "A reserved Windows-key combination was accepted.");

            var criticalSystem = ShortcutPolicy.Validate(
                new[] { "LeftAlt", "F4" },
                riskConfirmed: true);
            Assert(!criticalSystem.IsValid &&
                   criticalSystem.Error.IndexOf("critical", StringComparison.OrdinalIgnoreCase) >= 0,
                "A chord reserved for critical Windows or PowerPoint behavior was accepted.");

            var arbitraryLength = ShortcutPolicy.Validate(
                new[] { "LeftCtrl", "LeftAlt", "LeftShift", "F6", "G", "H", "J" },
                riskConfirmed: true);
            Assert(arbitraryLength.IsValid && arbitraryLength.Chord.Keys.Count == 7,
                "An arbitrary supported chord length was rejected.");

            var existing = new[]
            {
                new ShortcutAssignmentState
                {
                    FeatureId = "ExistingCommand",
                    Chord = normalized.Chord,
                    IsActive = true
                }
            };
            var conflict = ShortcutPolicy.Validate(
                new[] { "LeftCtrl", "LeftAlt", "M" },
                "ReplacementCommand",
                existing,
                riskConfirmed: true);
            Assert(!conflict.IsValid &&
                   conflict.ConflictingFeatureId == "ExistingCommand",
                "A duplicate openEfficiency assignment was accepted.");

            AssertRiskConfirmationRequired(new[] { "F8" }, "single-key");
            AssertRiskConfirmationRequired(new[] { "LeftCtrl", "LeftAlt" }, "modifier-only");
            AssertRiskConfirmationRequired(new[] { "G", "A" }, "unmodified text-key");

            Assert(!ShortcutActivationPolicy.CanExecute(
                    false, true, true, true, false),
                "A shortcut can execute while PowerPoint is not foreground.");
            Assert(!ShortcutActivationPolicy.CanExecute(
                    true, false, true, true, false),
                "A shortcut can execute without an eligible presentation window.");
            Assert(!ShortcutActivationPolicy.CanExecute(
                    true, true, false, true, false),
                "A shortcut can execute before add-in initialization.");
            Assert(!ShortcutActivationPolicy.CanExecute(
                    true, true, true, false, false),
                "A shortcut can execute before command-handler initialization.");
            Assert(!ShortcutActivationPolicy.CanExecute(
                    true, true, true, true, true),
                "Modal chord capture does not prevent command execution.");
            Assert(ShortcutActivationPolicy.CanExecute(
                    true, true, true, true, false),
                "A fully eligible foreground shortcut cannot execute.");

            string active = "Ctrl + A";
            string persisted = active;
            bool replacementResult = ShortcutReplacementTransaction.TryReplace(
                active,
                "Ctrl + B",
                replacement => false,
                replacement => persisted = replacement,
                previous => active = previous,
                out string replacementError);
            Assert(!replacementResult &&
                   active == "Ctrl + A" &&
                   persisted == "Ctrl + A" &&
                   !string.IsNullOrWhiteSpace(replacementError),
                "Failed replacement did not preserve the previous active and persisted assignment.");

            var migrated = ShortcutPolicy.ClassifyLegacyAssignments(new[]
            {
                new KeyValuePair<string, string>(
                    "LegacyValid",
                    "LeftCtrl,LeftAlt,M,None"),
                new KeyValuePair<string, string>(
                    "LegacyInvalid",
                    "LeftCtrl,RightCtrl,M,None")
            });
            var validLegacy = migrated.Single(state => state.FeatureId == "LegacyValid");
            var invalidLegacy = migrated.Single(state => state.FeatureId == "LegacyInvalid");
            Assert(validLegacy.IsActive &&
                   validLegacy.ShortcutText == "Ctrl + Alt + M",
                "A valid legacy shortcut was not migrated.");
            Assert(!invalidLegacy.IsActive &&
                   invalidLegacy.NeedsAttention &&
                   !string.IsNullOrWhiteSpace(invalidLegacy.ShortcutText),
                "An invalid legacy shortcut was not retained visibly as inactive.");

            var conflictingLegacy = ShortcutPolicy.ClassifyLegacyAssignments(new[]
            {
                new KeyValuePair<string, string>("LegacyOne", "LeftCtrl,K"),
                new KeyValuePair<string, string>("LegacyTwo", "RightCtrl,K")
            });
            Assert(conflictingLegacy.All(state => !state.IsActive && state.NeedsAttention),
                "Conflicting legacy assignments were registered.");

            Assert(!Directory.GetFiles(root, "shortcuts.cfg", SearchOption.AllDirectories).Any(),
                "The shipped default shortcut set is not empty.");

            var shortcutManagerSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Services", "ShortcutManager.cs"));
            var editorSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Dialogs", "ShortcutConfigWindow.xaml.cs"));
            var settingsSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "SettingsViewModel.cs"));
            var explorerSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Dialogs", "FeatureExplorerWindow.xaml.cs"));
            Assert(shortcutManagerSource.Contains("ShortcutActivationPolicy.CanExecute") &&
                   shortcutManagerSource.Contains("IsPowerPointForeground()"),
                "The runtime is not wired through the tested foreground activation policy.");
            Assert(editorSource.Contains("SuspendActivationForModalCapture"),
                "The shortcut editor does not suspend command activation while capturing a chord.");
            Assert(shortcutManagerSource.Contains("ShortcutReplacementTransaction.TryReplace"),
                "Shortcut replacement is not wired through the tested transaction.");
            Assert(settingsSource.Contains("assignment?.ShortcutText") &&
                   explorerSource.Contains("assignment.ShortcutText"),
                "Settings and Feature Explorer do not use the same normalized shortcut text.");
        }

        private static void AssertRiskConfirmationRequired(
            IEnumerable<string> keys,
            string scope)
        {
            var warning = ShortcutPolicy.Validate(keys);
            Assert(!warning.IsValid && warning.RequiresRiskConfirmation,
                $"The {scope} shortcut did not require explicit risk confirmation.");
            var confirmed = ShortcutPolicy.Validate(keys, riskConfirmed: true);
            Assert(confirmed.IsValid,
                $"The confirmed {scope} shortcut was not accepted.");
        }

        private static void ValidateRibbonAndDocumentationTruth(
            string root,
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            string powerPointRibbon = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "MainRibbon.xml"));
            string powerPointCallbacks = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "MainRibbon.cs"));
            string excelRibbon = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency.Excel", "UI", "ExcelRibbon.xml"));
            string excelRibbonCallbacks = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency.Excel", "UI", "ExcelRibbon.cs"));
            string sidebarSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "MainSidebar.xaml.cs")) +
                Environment.NewLine + File.ReadAllText(Path.Combine(
                    root, "oPenEfficiency", "UI", "Panels", "MainSidebar.ContextMenus.cs")) +
                Environment.NewLine + File.ReadAllText(Path.Combine(
                    root, "oPenEfficiency", "UI", "Panels", "MainSidebar.ContextMenus.Layout.cs")) +
                Environment.NewLine + File.ReadAllText(Path.Combine(
                    root, "oPenEfficiency", "UI", "Panels", "MainSidebar.ContextMenus.WorkflowAndVisual.cs")) +
                Environment.NewLine + File.ReadAllText(Path.Combine(
                    root, "oPenEfficiency", "UI", "Panels", "MainSidebar.ContextMenus.TableAndObject.cs")) +
                Environment.NewLine + File.ReadAllText(Path.Combine(
                    root, "oPenEfficiency", "UI", "Panels", "MainSidebar.ContextMenus.TextFormatting.cs")) +
                Environment.NewLine + File.ReadAllText(Path.Combine(
                    root, "oPenEfficiency", "UI", "Panels", "MainSidebar.RemainingContextMenus.cs"));
            string explorerSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Dialogs", "FeatureExplorerWindow.xaml.cs"));
            string explorerXaml = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Dialogs", "FeatureExplorerWindow.xaml"));
            string settingsSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "SettingsViewModel.cs"));
            string readme = File.ReadAllText(Path.Combine(root, "README.md"));
            string featureArchitecture = File.ReadAllText(Path.Combine(
                root, "docs", "development", "features.md"));

            Assert(powerPointRibbon.Contains("<tab idMso=\"TabHome\">"),
                "The PowerPoint ribbon launcher is not on the Home tab.");
            Assert(!powerPointRibbon.Contains("insertAfterMso=\"GroupEditingExcel\""),
                "The PowerPoint launcher uses an Excel-only insertion anchor and will not render.");
            Assert(Regex.IsMatch(
                    powerPointRibbon,
                    @"<button\s+id=""ToggleButtonSidebar""(?:(?!</?button).)*label=""openEfficiency""",
                    SourceOptions),
                "The PowerPoint Sidebar launcher does not use the canonical label.");
            Assert(powerPointRibbon.Contains("screentip=\"Show/hide openEfficiency Sidebar\"") &&
                   powerPointRibbon.Contains("supertip=\"Shows or hides the openEfficiency Sidebar for the current PowerPoint window.\""),
                "The PowerPoint launcher does not use the canonical action-and-consequence hover hierarchy.");
            Assert(powerPointRibbon.Contains("onAction=\"OnToggleSidebar\"") &&
                   powerPointCallbacks.Contains("Globals.ThisAddIn.ToggleSidebar()"),
                "The PowerPoint launcher is not wired to the Sidebar visibility toggle.");
            Assert(!Regex.IsMatch(
                    powerPointRibbon + Environment.NewLine + powerPointCallbacks,
                    @"\bopenE\b",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant),
                "Legacy 'openE' remains in the user-facing PowerPoint ribbon or callbacks.");

            Assert(Regex.IsMatch(
                    excelRibbon,
                    @"<tab\s+id=""TabOPenEfficiency""\s+label=""openEfficiency"">",
                    RegexOptions.CultureInvariant),
                "The optional Excel companion does not retain its dedicated openEfficiency tab.");

            // The build marker answers "which build is loaded?" in the ribbon, so its
            // callback wiring is pinned with the same kind of check as the launcher above.
            Assert(powerPointRibbon.Contains("getLabel=\"GetBuildStampLabel\"") &&
                   powerPointCallbacks.Contains("public string GetBuildStampLabel(IRibbonControl control)"),
                "The PowerPoint ribbon build marker is not wired to its label callback.");
            Assert(excelRibbon.Contains("getLabel=\"GetBuildStampLabel\"") &&
                   excelRibbonCallbacks.Contains("public string GetBuildStampLabel(IRibbonControl control)"),
                "The Excel ribbon build marker is not wired to its label callback.");

            var variantRecords = catalog
                .Where(record => record.Variants != null && record.Variants.Count > 0)
                .ToList();
            foreach (var record in variantRecords)
            {
                foreach (string variant in record.Variants)
                {
                    Assert(record.DetailedHelpText.IndexOf(
                            variant,
                            StringComparison.OrdinalIgnoreCase) >= 0,
                        $"Catalog variant '{record.Id}: {variant}' is absent from Detailed Help.");
                }

                AssertVariantMenuUsesCatalogLabels(sidebarSource, record);
            }

            string sidebarFactory = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "SidebarControlFactory.cs"));
            Assert(sidebarFactory.Contains("Right-click for more options") &&
                   sidebarFactory.Contains("info.Variants != null"),
                "Sidebar hover cards do not generate the standard context-menu capability hint.");

            string readableHelp = FeatureHelpFormatter.ToReadablePlainText(
                "### Sample" + Environment.NewLine +
                "**Usage:**" + Environment.NewLine +
                "- First option");
            Assert(readableHelp == "Sample" + Environment.NewLine +
                   "Usage:" + Environment.NewLine + "• First option",
                "Markdown-like Detailed Help is not converted to readable plain text.");
            Assert(explorerSource.Contains("FeatureHelpFormatter.ToReadablePlainText") &&
                   explorerSource.Contains("FeatureCatalogSearch.Matches") &&
                   explorerSource.Contains(".Where(record => record.UserVisible && !record.IsObsolete)") &&
                   explorerSource.Contains("record.Category") &&
                   explorerSource.Contains("record.Maturity") &&
                   explorerSource.Contains("assignment.ShortcutText") &&
                   explorerSource.Contains("if (!string.IsNullOrEmpty(feature.HelpImagePath))"),
                "Feature Explorer is not using catalog categories, aliases, readable help, maturity, and normalized shortcut text.");
            foreach (string category in new[]
            {
                "Align &amp; Size", "Arrange &amp; Organize", "Text &amp; Format",
                "Tables", "Visual Elements", "Storyline &amp; Review",
                "Data &amp; Automation", "Libraries", "Settings &amp; Help", "Labs"
            })
            {
                Assert(explorerXaml.Contains("Tag=\"" + category + "\""),
                    $"Feature Explorer has no usable filter for catalog category '{category}'.");
            }
            Assert(settingsSource.Contains("FeatureCatalogSearch.Matches") &&
                   sidebarSource.Contains("SidebarSearchService.Search"),
                "Settings, Command Search, and Feature Explorer are not sharing catalog alias search.");

            Assert(readme.Contains("PowerPoint **Home** tab") &&
                   readme.Contains("primary PowerPoint command surface"),
                "README does not accurately describe the Home-tab Sidebar launcher.");
            Assert(readme.Contains("**Consulting Standard**") &&
                   readme.Contains("**All Features**") &&
                   readme.Contains("**Minimal**") &&
                   readme.Contains("**Labs**"),
                "README does not list the four shipped preset names.");
            Assert(readme.Contains("ships with no default feature shortcuts"),
                "README does not preserve the no-default-shortcuts statement.");
            Assert(readme.Contains("optional Excel companion") &&
                   readme.Contains("dedicated **openEfficiency** Excel ribbon tab") &&
                   readme.Contains("intentionally uses the Home-tab Sidebar launcher"),
                "README does not explain the optional Excel companion and its different primary surface.");
            Assert(readme.Contains("**Preview:**") &&
                   readme.Contains("**Experimental:**") &&
                   readme.Contains("native PowerPoint shapes"),
                "README does not truthfully describe maturity and native-shape positioning.");
            Assert(!readme.Contains("custom \"oPen Efficiency\" tab") &&
                   !readme.Contains("**Minimalist**") &&
                   !readme.Contains("Ctrl+Alt+S") &&
                   !readme.Contains("Ctrl+Alt+M") &&
                   !readme.Contains("Ctrl+Alt+A"),
                "README retains a contradicted ribbon, preset, or default-shortcut claim.");
            Assert(featureArchitecture.Contains(
                    "required source for every future user-facing command surface"),
                "Developer feature architecture does not make the catalog requirement explicit.");
        }

        private static void AssertVariantMenuUsesCatalogLabels(
            string sidebarSource,
            FeatureCatalogRecord record)
        {
            if (record.Id == "BtnAlignLeft" ||
                record.Id == "BtnAlignRight" ||
                record.Id == "BtnAlignTop" ||
                record.Id == "BtnAlignBottom")
            {
                string sharedStart =
                    "if (btn.Name == \"BtnAlignLeft\" || btn.Name == \"BtnAlignRight\"";
                int start = sidebarSource.IndexOf(sharedStart, StringComparison.Ordinal);
                int end = sidebarSource.IndexOf(
                    "else if (btn.Name == \"BtnSwapPositions\")",
                    start,
                    StringComparison.Ordinal);
                Assert(start >= 0 && end > start &&
                       CountOccurrences(
                           sidebarSource.Substring(start, end - start),
                           "GetCatalogVariantLabel") >= 4,
                    "Alignment context-menu variants are not bound to catalog labels.");
                return;
            }

            var markerFamilies = new[]
            {
                new
                {
                    Marker = "if (btn.Name == \"" + record.Id + "\")",
                    NextMarker = "else if (btn.Name =="
                },
                new
                {
                    Marker = "else if (btn.Name == \"" + record.Id + "\")",
                    NextMarker = "else if (btn.Name =="
                },
                new
                {
                    Marker = "if (featureId == \"" + record.Id + "\")",
                    NextMarker = "else if (featureId =="
                },
                new
                {
                    Marker = "else if (featureId == \"" + record.Id + "\")",
                    NextMarker = "else if (featureId =="
                },
                new
                {
                    Marker = "case \"" + record.Id + "\":",
                    NextMarker = "case \""
                }
            };

            int candidateCount = 0;
            int bestCatalogLabelCount = 0;
            int bestRemainingMenuItemCount = 0;
            foreach (var family in markerFamilies)
            {
                int searchStart = 0;
                while (searchStart < sidebarSource.Length)
                {
                    int blockStart = sidebarSource.IndexOf(
                        family.Marker,
                        searchStart,
                        StringComparison.Ordinal);
                    if (blockStart < 0) break;

                    candidateCount++;
                    int blockEnd = sidebarSource.IndexOf(
                        family.NextMarker,
                        blockStart + family.Marker.Length,
                        StringComparison.Ordinal);
                    if (blockEnd < 0) blockEnd = sidebarSource.Length;
                    string block = sidebarSource.Substring(blockStart, blockEnd - blockStart);
                    int catalogLabelCount = CountOccurrences(block, "GetCatalogVariantLabel");
                    int remainingMenuItemCount = CountOccurrences(block, "AddRemainingMenuItem");
                    bestCatalogLabelCount = Math.Max(bestCatalogLabelCount, catalogLabelCount);
                    bestRemainingMenuItemCount = Math.Max(
                        bestRemainingMenuItemCount,
                        remainingMenuItemCount);

                    if (catalogLabelCount >= record.Variants.Count ||
                        remainingMenuItemCount >= record.Variants.Count)
                    {
                        return;
                    }

                    searchStart = blockStart + family.Marker.Length;
                }
            }

            Assert(candidateCount > 0,
                $"Catalog command '{record.Id}' has no Sidebar context-menu block.");
            Assert(false,
                $"Context-menu labels for '{record.Id}' are not fully sourced from catalog variants " +
                $"(variants: {record.Variants.Count}, best catalog labels: {bestCatalogLabelCount}, " +
                $"best remaining menu items: {bestRemainingMenuItemCount}, " +
                $"candidate blocks: {candidateCount}).");
        }

        /// <summary>
        /// BtnSearchBar behavioral contract: pressing Enter runs GetSearchMatches().FirstOrDefault(),
        /// so the ordering ladder decides which command the user actually launches. This pins the full
        /// production scoring ladder of FeatureCatalogSearch.GetMatchScore (exact name/id -> exact
        /// alias -> name prefix -> alias prefix -> id prefix -> name contains -> alias contains -> id
        /// contains -> description -> help text, case-insensitive), the ordinal-ignore-case name
        /// tie-break on equal scores, and that obsolete records never enter the result.
        /// </summary>
        private static void ValidateSearchEnterPriorityContract()
        {
            FeatureCatalogRecord Record(string id, string name,
                string alias = null, string description = null, string help = null, bool obsolete = false)
            {
                return new FeatureCatalogRecord
                {
                    Id = id,
                    Name = name,
                    Description = description,
                    DetailedHelpText = help,
                    SearchAliases = alias == null ? Array.Empty<string>() : new[] { alias },
                    UserVisible = true,
                    IsObsolete = obsolete
                };
            }

            var catalog = new[]
            {
                Record("help-holder", "Help Holder", help: "mentions Anchor"),
                Record("desc-holder", "Desc Holder", description: "find Anchor here"),
                Record("zulu-anchor", "Zulu Anchor"),
                Record("alias-contains", "Alias Holder", alias: "the anchor"),
                Record("XANCHORX", "Id Contain"),
                Record("AnchorId", "Id Only"),
                Record("anchor-exact", "Anchor"),
                Record("name-prefix", "Anchor Prime"),
                Record("alias-prefix", "Prefix Holder", alias: "Anchor X"),
                Record("name-contains-a", "The Anchor Tool"),
                Record("alias-exact", "Other Tool", alias: "Anchor"),
                Record("alpha-anchor", "alpha anchor"),
                Record("obsolete-anchor", "Anchor", obsolete: true),
            };
            var matches = SidebarSearchService.Search(catalog, "Anchor");
            string ordered = string.Join("|", matches.Select(match => match.Id));

            TestAssert.AreEqual(
                "anchor-exact|alias-exact|name-prefix|alias-prefix|AnchorId|alpha-anchor|name-contains-a|zulu-anchor|alias-contains|XANCHORX|desc-holder|help-holder",
                ordered,
                "The sidebar search priority ladder, tie-break or obsolete filtering changed - Enter would launch a different command.");

            TestAssert.AreEqual(0, SidebarSearchService.Search(catalog, "   ").Count,
                "An empty or whitespace query must return no results (Enter must not launch anything).");
        }

        /// <summary>
        /// BtnQuickContextMenu crash guard: QuickContextMenuService builds its popup dictionary with
        /// catalog[id] for every configured id, so a renamed, hidden or obsolete default id would
        /// throw a KeyNotFoundException at popup time. This pins that every default id still resolves
        /// as a user-visible, non-obsolete command in the current catalog.
        /// </summary>
        private static void ValidateQuickContextMenuDefaults(IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            var byId = catalog.ToDictionary(record => record.Id, StringComparer.Ordinal);
            foreach (string id in oPenEfficiency.Models.QuickContextMenuSettings.DefaultFeatureIds)
            {
                Assert(byId.TryGetValue(id, out FeatureCatalogRecord record),
                    $"Quick Context Menu default '{id}' no longer exists in the catalog - the popup would throw KeyNotFoundException.");
                Assert(record.UserVisible && !record.IsObsolete &&
                       record.ExecutionKind == FeatureExecutionKind.Command,
                    $"Quick Context Menu default '{id}' is not a user-visible command - the popup filters it out or crashes.");
            }
        }

        private static void ValidateSidebarUiLogic()
        {
            var catalog = new[]
            {
                new FeatureCatalogRecord { Id = "contains", Name = "Alpha Tool", Description = "Needle", UserVisible = true },
                new FeatureCatalogRecord { Id = "exact", Name = "Needle", UserVisible = true },
                new FeatureCatalogRecord { Id = "widget", Name = "Needle Widget", UserVisible = true, ExecutionKind = FeatureExecutionKind.Widget },
                new FeatureCatalogRecord { Id = "hidden", Name = "Needle Hidden", UserVisible = false }
            };
            var matches = SidebarSearchService.Search(catalog, "needle");
            Assert(matches.Count == 2 && matches[0].Id == "exact" && matches[1].Id == "contains",
                "Sidebar search priority or command filtering changed.");
            Assert(SidebarSearchService.Search(catalog, "  ").Count == 0,
                "An empty sidebar query returned results.");

            var rule = new SidebarFeatureStateRule
            {
                MinSelection = 2,
                MaxSelection = 3,
                RequiredType = OfficeSelectionType.Shapes
            };
            Assert(SidebarStateEvaluator.IsEnabled(rule, new SidebarSelectionSnapshot
                { Type = OfficeSelectionType.Shapes, ShapeCount = 2 }),
                "A valid multi-shape selection was disabled.");
            Assert(!SidebarStateEvaluator.IsEnabled(rule, new SidebarSelectionSnapshot
                { Type = OfficeSelectionType.Shapes, ShapeCount = 1 }),
                "The minimum selection rule was ignored.");
            Assert(!SidebarStateEvaluator.IsEnabled(rule, new SidebarSelectionSnapshot
                { Type = OfficeSelectionType.Shapes, ShapeCount = 4 }),
                "The maximum selection rule was ignored.");

            var tableRule = new SidebarFeatureStateRule { RequiresTable = true };
            Assert(SidebarStateEvaluator.IsEnabled(tableRule, new SidebarSelectionSnapshot
                { Type = OfficeSelectionType.Text, HasTable = true }),
                "A text cursor inside a table was disabled.");
            Assert(!SidebarStateEvaluator.IsEnabled(tableRule, new SidebarSelectionSnapshot
                { Type = OfficeSelectionType.Shapes }),
                "A table-only feature was enabled without a table.");

            var hiddenRule = new SidebarFeatureStateRule { RequiresHiddenShapes = true };
            Assert(SidebarStateEvaluator.IsEnabled(hiddenRule, new SidebarSelectionSnapshot
                { HasHiddenShapes = true }), "Hidden-shape availability was ignored.");
            Assert(!SidebarStateEvaluator.IsEnabled(hiddenRule, new SidebarSelectionSnapshot()),
                "A hidden-shape feature was enabled without hidden shapes.");

            Assert(SidebarStateEvaluator.HasUsableTextTarget("BtnApplyTextTool",
                new SidebarSelectionSnapshot
                {
                    Type = OfficeSelectionType.Shapes,
                    ShapeCount = 1,
                    TextOrTableShapeCount = 1,
                    HasTable = true
                }), "Apply Text rejected a selected table.");
            Assert(SidebarStateEvaluator.HasUsableTextTarget("BtnSplitByParagraphs",
                new SidebarSelectionSnapshot
                {
                    Type = OfficeSelectionType.Text,
                    HasTextRange = true
                }), "Split Text Box by Paragraphs rejected a text selection.");

            Assert(SidebarStateEvaluator.ResolveAggregateToggleState(true, 2, 2) == true,
                "A fully active toggle selection was not checked.");
            Assert(SidebarStateEvaluator.ResolveAggregateToggleState(true, 2, 1) == null,
                "A mixed toggle selection was not indeterminate.");
            Assert(SidebarStateEvaluator.ResolveAggregateToggleState(false, 2, 2) == false,
                "A disabled toggle was left checked.");
        }

        private static readonly (string Word, int Value)[] ShapeCountWords =
        {
            ("two", 2), ("three", 3), ("four", 4), ("five", 5),
            ("six", 6), ("seven", 7), ("eight", 8), ("nine", 9), ("ten", 10)
        };

        /// <summary>
        /// Review round 6c measured that the third batch asserted only metadata -> SidebarStateEvaluator:
        /// rewriting the distribute commands' own selection guard left all eleven scenarios green. This
        /// assertion binds the command's guard to the host-neutral production rule that both Execute methods
        /// must use, and the rule's numeric contract is executed by DistributeSelectionGuardTests (scenarios
        /// QA9-BtnDistributeHorizontal-002 / -Vertical-002). The COM effect stays the office-automated leg.
        /// </summary>
        private static void ValidateDistributeSelectionGuardDelegation(string repositoryRoot, IReadOnlyList<FeatureCatalogRecord> catalog, string featureId)
        {
            var record = catalog.FirstOrDefault(item => string.Equals(item.Id, featureId, StringComparison.Ordinal));
            Assert(record != null, $"{featureId} is missing from the production catalog, so its selection guard cannot be accepted.");
            Assert(record.MinSelection == 2, $"{featureId} no longer declares MinSelection=2, so the documented two-shape minimum changed.");

            var featureSource = File.ReadAllText(Path.Combine(repositoryRoot, "oPenEfficiency", "Features", "Alignment", "DistributeFeatures.cs"));
            var planCalls = Regex.Matches(featureSource, @"DistributeSelectionRule\.ResolvePlan\(\s*sr\.Count\s*,\s*relativeToSlide\s*\)").Count;
            Assert(planCalls == 2, $"{featureId}: both Execute methods must resolve the whole decision through the production rule (found {planCalls}) - without it the enabled click at two shapes is a no-op again (REG-DISTRIBUTE-SCOPE-001).");
            var planGuards = Regex.Matches(featureSource, @"if \(!plan\.Distribute\) return false;").Count;
            Assert(planGuards == 2, $"{featureId}: the plan's Distribute decision must be the command's guard (found {planGuards}).");
            var routeUsage = Regex.Matches(featureSource, @"plan\.RelativeToSlide \? Office\.MsoTriState\.msoTrue : Office\.MsoTriState\.msoFalse").Count;
            Assert(routeUsage == 2, $"{featureId}: the resolved route must be the value the distribution receives, otherwise the rule is consulted and ignored (found {routeUsage}).");
            var planAssignments = Regex.Matches(featureSource, @"\bplan\s*=(?!=)").Count;
            Assert(planAssignments == 2, $"{featureId}: the plan must be resolved exactly once in each Execute method (found {planAssignments} assignments) - a second assignment can flip both decisions while every rule call stays verbatim.");
            var parameterAssignments = Regex.Matches(featureSource, @"\brelativeToSlide\s*=(?!=)").Count;
            Assert(parameterAssignments == 0, $"{featureId}: relativeToSlide is assigned inside the command (found {parameterAssignments}) - flipping it inverts the minimum each route applies while every call stays verbatim (review round 6f).");
            // Review round 6f: the previous inline check only looked at `sr.Count` followed by a comparison sign in
            // this one file. Measured defeaters that stayed green: the count comparison moved into a helper method of
            // the same file (`Refuse(int count) { return count < 5; }` OR-ed into both guards), the same helper in a
            // second file, and a second guard written as `sr.Count == 5`. The rule is class-wide now - every product
            // file that takes part in the distribute guard must be free of count-versus-number comparisons - and the
            // route decision must be computed exactly once in each Execute method.
            var participants = new List<string>
            {
                Path.Combine(repositoryRoot, "oPenEfficiency", "Features", "Alignment", "DistributeFeatures.cs"),
                Path.Combine(repositoryRoot, "oPenEfficiency.Core", "Services", "DistributeSelectionRule.cs"),
            };
            // Review round 6g: the scan root was oPenEfficiency/ only, so a guard helper under oPenEfficiency.Core/
            // or any other product project stayed outside the participant set. Every product source tree is scanned
            // now; participation is still decided by the file name or by mentioning the rule, so unrelated files
            // stay out.
            var scanRoots = new[]
            {
                Path.Combine(repositoryRoot, "oPenEfficiency"),
                Path.Combine(repositoryRoot, "oPenEfficiency.Core"),
            };
            var participantCandidates = new List<string>();
            foreach (var scanRoot in scanRoots)
            {
                if (Directory.Exists(scanRoot))
                {
                    participantCandidates.AddRange(Directory.GetFiles(scanRoot, "*.cs", SearchOption.AllDirectories));
                }
            }

            foreach (var candidate in participantCandidates)
            {
                if (participants.Contains(candidate, StringComparer.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (Path.GetFileName(candidate).StartsWith("Distribute", StringComparison.Ordinal)
                    || File.ReadAllText(candidate).IndexOf("DistributeSelectionRule", StringComparison.Ordinal) >= 0)
                {
                    participants.Add(candidate);
                }
            }

            foreach (var participant in participants)
            {
                var source = File.ReadAllText(participant);
                var label = participant.Substring(repositoryRoot.Length).TrimStart(Path.DirectorySeparatorChar);
                Assert(
                    !Regex.IsMatch(source, @"\b[\w.]*[Cc]ount[\w.]*\s*(?:[<>]=?|[=!]=)\s*\d"),
                    $"{featureId}: '{label}' compares a shape count against a number instead of leaving the decision to the production rule (review round 6f measured a helper method, a second file and an '==' guard that all stayed green).");
            }

            var ruleSource = File.ReadAllText(Path.Combine(repositoryRoot, "oPenEfficiency.Core", "Services", "DistributeSelectionRule.cs"));
            Assert(
                Regex.IsMatch(ruleSource, @"ResolvePlan\(int selectedShapeCount, bool relativeToSlide\)"),
                "The production rule no longer resolves the whole distribute decision in one call (review round 6g).");
            Assert(
                Regex.IsMatch(ruleSource, @"UsesAcrossSlideRoute\(int selectedShapeCount, bool relativeToSlide\)"),
                $"{featureId}: the production rule no longer carries the route decision the two-shape fallback depends on (REG-DISTRIBUTE-SCOPE-001).");
            Assert(
                Regex.IsMatch(ruleSource, @"MinimumShapesForPrimaryRoute\s*=\s*3"),
                "The primary route's minimum of three shapes is the measured contract; changing it needs the product decision recorded in REG-DISTRIBUTE-SCOPE-001.");
            Assert(
                Regex.IsMatch(ruleSource, @"MinimumShapesForAcrossSlideRoute\s*=\s*2"),
                "The across-slide route's minimum of two shapes is what the metadata minimum promises.");
        }

        /// <summary>
        /// Wave 1 (QA-S9) selection-scope contract: the enabled/disabled decision a feature's own metadata
        /// produces through the production evaluator. The scenario asserts the documented promise - the
        /// documented minimum is usable, anything below it is refused, an out-of-scope selection is refused -
        /// and not merely that the numbers are present, so a weakened MinSelection or a changed evaluator
        /// fails here. The COM effect of these commands stays the office-automated leg.
        /// </summary>
        private static void ValidateWave1SelectionScopeContract(IReadOnlyList<FeatureCatalogRecord> catalog, string featureId)
        {
            var record = catalog.FirstOrDefault(item => string.Equals(item.Id, featureId, StringComparison.Ordinal));
            Assert(record != null, $"{featureId} is missing from the production catalog, so its selection scope cannot be accepted.");
            Assert(record.UserVisible, $"{featureId} is no longer user-visible; a selection-scope scenario only covers a live command.");

            var rule = new SidebarFeatureStateRule
            {
                MinSelection = record.MinSelection,
                MaxSelection = record.MaxSelection,
                RequiredType = record.RequiredType,
                RequiresTable = record.RequiresTable,
                RequiresHiddenShapes = record.RequiresHiddenShapes
            };

            bool expectedWithoutSelection = record.MinSelection == 0 && !record.RequiresTable && !record.RequiresHiddenShapes;
            Assert(
                SidebarStateEvaluator.IsEnabled(rule, null) == expectedWithoutSelection,
                $"{featureId} enables on an empty selection although its metadata (MinSelection={record.MinSelection}, " +
                $"RequiresTable={record.RequiresTable}, RequiresHiddenShapes={record.RequiresHiddenShapes}) does not allow it.");

            if (record.MinSelection > 0)
            {
                Assert(
                    !SidebarStateEvaluator.IsEnabled(rule, ShapesSelection(record.MinSelection - 1)),
                    $"{featureId} was enabled below its documented minimum of {record.MinSelection} selected shape(s).");
                Assert(
                    SidebarStateEvaluator.IsEnabled(rule, ShapesSelection(record.MinSelection)),
                    $"{featureId} was refused at its documented minimum of {record.MinSelection} selected shape(s).");
                if (record.RequiredType == OfficeSelectionType.Shapes)
                {
                    Assert(
                        !SidebarStateEvaluator.IsEnabled(rule, new SidebarSelectionSnapshot { Type = OfficeSelectionType.Slides }),
                        $"{featureId} is enabled for a slide selection although it requires selected shapes.");
                }
            }
            else
            {
                Assert(
                    SidebarStateEvaluator.IsEnabled(rule, ShapesSelection(1)),
                    $"{featureId} declares no minimum selection but refused a single selected shape.");
            }

            if (record.MaxSelection > 0)
            {
                Assert(
                    !SidebarStateEvaluator.IsEnabled(rule, ShapesSelection(record.MaxSelection + 1)),
                    $"{featureId} was enabled above its documented maximum of {record.MaxSelection} selected shape(s).");
            }

            var documentedMinimum = ReadLowestDocumentedShapeMinimum(record);
            if (documentedMinimum.HasValue)
            {
                Assert(
                    record.MinSelection <= documentedMinimum.Value,
                    $"{featureId} documents {documentedMinimum.Value} or more shapes but requires " +
                    $"{record.MinSelection} selected shape(s), so the documented route cannot be used.");
                Assert(
                    SidebarStateEvaluator.IsEnabled(rule, ShapesSelection(documentedMinimum.Value)),
                    $"{featureId} documents {documentedMinimum.Value} or more shapes yet was refused at that scope.");
            }
        }

        private static int? ReadLowestDocumentedShapeMinimum(FeatureCatalogRecord record)
        {
            string text = string.Join(" ", new[] { record.Tooltip, record.Description, record.DetailedHelpText });
            if (string.IsNullOrEmpty(text)) return null;

            var minima = new List<int>();
            foreach (Match match in Regex.Matches(text, @"(\d+)\s+or more shapes", RegexOptions.CultureInvariant))
            {
                minima.Add(int.Parse(match.Groups[1].Value));
            }
            foreach (var word in ShapeCountWords)
            {
                if (Regex.IsMatch(text, word.Word + @"\s+or more shapes", RegexOptions.CultureInvariant))
                {
                    minima.Add(word.Value);
                }
            }
            return minima.Count == 0 ? (int?)null : minima.Min();
        }

        private static SidebarSelectionSnapshot ShapesSelection(int count)
        {
            return new SidebarSelectionSnapshot
            {
                Type = OfficeSelectionType.Shapes,
                ShapeCount = count,
                TextShapeCount = count,
                TextOrTableShapeCount = count
            };
        }

        private static void ValidateLifecycleArchitecture(string root)
        {
            string addIn = File.ReadAllText(Path.Combine(root, "oPenEfficiency", "ThisAddIn.cs"));
            string gantt = File.ReadAllText(Path.Combine(root, "oPenEfficiency", "ThisAddIn.GanttUi.cs"));
            string modeless = File.ReadAllText(Path.Combine(root, "oPenEfficiency", "ThisAddIn.ModelessUi.cs"));
            string sidebar = string.Join(Environment.NewLine,
                Directory.GetFiles(Path.Combine(root, "oPenEfficiency", "UI", "Panels"), "MainSidebar*.cs")
                    .Select(File.ReadAllText));
            string exportWizard = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Wizards", "ExportWizard.xaml.cs"));

            Assert(addIn.Contains("SelectionUpdateCoordinator") &&
                   addIn.Contains("SidebarPaneCoordinator") &&
                   addIn.Contains("_selectionUpdates.Dispose()") &&
                   addIn.Contains("_sidebarPanes?.Dispose()"),
                "Sidebar pane or selection lifecycle is no longer coordinator-owned.");
            Assert(addIn.Contains("WindowSelectionChange += Application_WindowSelectionChange") &&
                   addIn.Contains("WindowSelectionChange -= Application_WindowSelectionChange") &&
                   addIn.Contains("_textEditPollingTimer.Tick -= TextEditPollingTimer_Tick"),
                "PowerPoint or timer event registration is no longer symmetric.");
            Assert(addIn.Contains("GanttInteractionCoordinator") &&
                   addIn.Contains("_ganttInteractionCoordinator.Dispose()") &&
                   gantt.Contains("_modelessUi.TrackWindow") &&
                   modeless.Contains("_modelessUi.TrackForm"),
                "Gantt or modeless UI cleanup is no longer centrally tracked.");
            Assert(!sidebar.Contains("Globals.ThisAddIn") &&
                   !exportWizard.Contains("Globals.ThisAddIn"),
                "Sidebar or Export Wizard regressed to the global add-in service locator.");

            int globalAccesses = Directory.GetFiles(
                    Path.Combine(root, "oPenEfficiency"), "*.cs", SearchOption.AllDirectories)
                .Where(path => path.IndexOf(
                        Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar,
                        StringComparison.OrdinalIgnoreCase) < 0 &&
                    path.IndexOf(
                        Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar,
                        StringComparison.OrdinalIgnoreCase) < 0)
                .Sum(path => CountOccurrences(File.ReadAllText(path), "Globals.ThisAddIn"));
            Assert(globalAccesses <= 25,
                $"Global add-in accesses increased from the Wave F ceiling: {globalAccesses}.");
        }

        private static int CountOccurrences(string source, string value)
        {
            int count = 0;
            int index = 0;
            while ((index = source.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
            {
                count++;
                index += value.Length;
            }
            return count;
        }

        private static List<FeatureCatalogRecord> ReadCurrentCatalog(string root)
        {
            var featureRoot = Path.Combine(root, "oPenEfficiency", "Features");
            var featureLibraryPath = Path.Combine(root, "oPenEfficiency", "UI", "FeatureLibrary.cs");
            var attributed = new List<FeatureCatalogRecord>();

            foreach (var path in Directory.GetFiles(featureRoot, "*.cs", SearchOption.AllDirectories))
            {
                var source = File.ReadAllText(path);
                foreach (Match match in Regex.Matches(
                    source,
                    @"\[FeatureMetadata\((.*?)\)\s*\]",
                    SourceOptions))
                {
                    var body = match.Groups[1].Value;
                    var id = ReadString(body, "Id");
                    if (string.IsNullOrWhiteSpace(id)) continue;

                    var record = new FeatureCatalogRecord
                    {
                        Id = id,
                        Name = ReadString(body, "Name"),
                        Tooltip = ReadString(body, "Tooltip"),
                        Description = ReadString(body, "Description"),
                        DetailedHelpText = ReadString(body, "DetailedHelpText"),
                        SearchAliases = ReadStringArray(body, "SearchAliases"),
                        MinSelection = ReadInt(body, "MinSelection"),
                        MaxSelection = ReadInt(body, "MaxSelection"),
                        RequiredType = ReadSelectionType(body),
                        RequiresTable = ReadBool(body, "RequiresTable"),
                        RequiresHiddenShapes = ReadBool(body, "RequiresHiddenShapes"),
                        IsToggle = ReadBool(body, "IsToggle"),
                        UserVisible = !HasExplicitFalse(body, "UserVisible")
                    };
                    FeatureCatalogPolicy.Apply(record);
                    attributed.Add(record);
                }
            }

            var attributedIds = new HashSet<string>(
                attributed.Select(record => record.Id),
                StringComparer.Ordinal);
            var librarySource = File.ReadAllText(featureLibraryPath);
            var manual = new List<FeatureCatalogRecord>();

            foreach (Match seed in Regex.Matches(
                librarySource,
                @"new\s+SidebarFeature\s*\{\s*Id\s*=\s*""([^""]+)"",\s*Name\s*=\s*""([^""]+)""\s*\}",
                SourceOptions))
            {
                var id = seed.Groups[1].Value;
                if (attributedIds.Contains(id) || manual.Any(record => record.Id == id)) continue;
                manual.Add(ReadManualRecord(librarySource, id, seed.Groups[2].Value));
            }

            if (!manual.Any(record => record.Id == "BtnExploreFeatures"))
                manual.Add(ReadManualRecord(librarySource, "BtnExploreFeatures", "Feature Explorer"));

            var shadowErrors = FeatureCatalogValidator.ValidateManualShadowing(
                attributedIds,
                manual);
            AssertNoErrors("manual catalog shadowing", shadowErrors);

            attributed.AddRange(manual);
            return attributed;
        }

        private static void ValidateCanonicalFeatureInventory(
            string root,
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            string path = Path.Combine(root, "docs", "product", "feature-catalog.md");
            Assert(File.Exists(path), "The canonical numbered feature catalog is missing.");

            var records = catalog
                .Where(record => record.UserVisible && !record.IsObsolete)
                .OrderBy(record => record.Category)
                .ThenBy(record => record.Name, StringComparer.Ordinal)
                .ThenBy(record => record.Id, StringComparer.Ordinal)
                .ToList();
            var expectedRows = records
                .Select((record, index) =>
                {
                    string type = record.ExecutionKind == FeatureExecutionKind.Widget
                        ? "Widget"
                        : record.IsToggle ? "Toggle" : "Command";
                    string registration = record.IsManual ? "Manual" : "Metadata";
                    string name = (record.Name ?? string.Empty).Replace("|", "\\|");
                    return $"| {index + 1} | `{record.Id}` | {name} | {record.Category} | {type} | {registration} | {record.Maturity} | {record.DefaultPlacement} |";
                })
                .ToList();
            var source = File.ReadAllText(path);
            var actualRows = File.ReadAllLines(path)
                .Where(line => Regex.IsMatch(line, @"^\|\s*\d+\s*\|\s*`Btn"))
                .ToList();

            Assert(actualRows.SequenceEqual(expectedRows),
                "The canonical numbered feature catalog does not match the resolved catalog. Regenerate docs/product/feature-catalog.md.");
            Assert(source.Contains($"Current total: **{records.Count} features**"),
                "The canonical feature catalog summary has an outdated total.");
        }

        private static void ValidateIconAccentAudit(
            string root,
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            string path = Path.Combine(root, "oPenEfficiency.Catalog.Tests", "TestData", "icon-accent-audit.json");
            List<IconAccentAuditRecord> audit;
            var serializer = new DataContractJsonSerializer(typeof(List<IconAccentAuditRecord>));
            using (var stream = File.OpenRead(path))
            {
                audit = (List<IconAccentAuditRecord>)serializer.ReadObject(stream);
            }

            var commandIds = catalog
                // Keep hidden compatibility commands in the icon contract while they remain
                // executable by a stable ID; only user-facing catalog inventories exclude them.
                .Where(record => !record.IsObsolete &&
                    record.ExecutionKind != FeatureExecutionKind.Widget)
                .Select(record => record.Id)
                .OrderBy(id => id, StringComparer.Ordinal)
                .ToList();
            var commandRows = audit
                .Where(record => string.Equals(record.Scope, "Command", StringComparison.Ordinal))
                .ToList();

            TestAssert.AreEqual(179, commandIds.Count, "The current command catalog count changed; document the legitimate change.");
            TestAssert.AreEqual(commandIds.Count, commandRows.Count, "The accent audit must contain exactly one row per command.");
            CollectionAssert.AreEqual(
                commandIds,
                commandRows.Select(record => record.FeatureId).OrderBy(id => id, StringComparer.Ordinal).ToList());
            TestAssert.AreEqual(
                commandRows.Count,
                commandRows.Select(record => record.FeatureId).Distinct(StringComparer.Ordinal).Count(),
                "The accent audit contains duplicate command rows.");
            TestAssert.IsTrue(audit
                .Where(record => !string.Equals(record.Scope, "Command", StringComparison.Ordinal))
                .All(record => string.Equals(record.Scope, "Extra", StringComparison.Ordinal)),
                "Non-command SVGs must be classified explicitly as Extra.");

            foreach (IconAccentAuditRecord record in commandRows)
            {
                TestAssert.IsFalse(string.IsNullOrWhiteSpace(record.Render16Light), $"{record.FeatureId} lacks the 16px Light result.");
                TestAssert.IsFalse(string.IsNullOrWhiteSpace(record.Render16Dark), $"{record.FeatureId} lacks the 16px Dark result.");
                TestAssert.IsFalse(string.IsNullOrWhiteSpace(record.Render24Light), $"{record.FeatureId} lacks the 24px Light result.");
                TestAssert.IsFalse(string.IsNullOrWhiteSpace(record.Render24Dark), $"{record.FeatureId} lacks the 24px Dark result.");
                TestAssert.IsFalse(string.IsNullOrWhiteSpace(record.Render48Light), $"{record.FeatureId} lacks the 48px Light result.");
                TestAssert.IsFalse(string.IsNullOrWhiteSpace(record.Render48Dark), $"{record.FeatureId} lacks the 48px Dark result.");
                TestAssert.IsFalse(string.IsNullOrWhiteSpace(record.ReviewerStatus), $"{record.FeatureId} lacks reviewer status.");
            }
        }

        private static void ValidateReviewerIconAccentContract(
            string root,
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            string[] reviewerPaths =
            {
                Path.Combine(root, "oPenEfficiency.Catalog.Tests", "TestData", "icon-accent-decisions-1-108.md"),
                Path.Combine(root, "oPenEfficiency.Catalog.Tests", "TestData", "icon-accent-decisions-109-176.md")
            };
            var decisions = new List<ReviewerIconAccentDecision>();
            var rowPattern = new Regex(
                @"^\|\s*\d+\s*\|\s*`(?<id>Btn[^`]+)`\s*\|\s*(?<decision>[^|]+?)\s*\|\s*[^|]*\|\s*(?<role>[^|]+?)\s*\|",
                RegexOptions.Multiline | RegexOptions.CultureInvariant);

            foreach (string reviewerPath in reviewerPaths)
            {
                TestAssert.IsTrue(File.Exists(reviewerPath), "Missing reviewer decision batch: " + reviewerPath);
                foreach (Match match in rowPattern.Matches(File.ReadAllText(reviewerPath)))
                {
                    string classification = match.Groups["decision"].Value.Trim();
                    bool isPartialAccent = classification.StartsWith("Partial accent", StringComparison.Ordinal);
                    bool isNeutral = classification.StartsWith("Neutral only", StringComparison.Ordinal);
                    TestAssert.IsTrue(isPartialAccent || isNeutral,
                        $"{match.Groups["id"].Value} has unsupported reviewer classification '{classification}'.");

                    decisions.Add(new ReviewerIconAccentDecision
                    {
                        FeatureId = match.Groups["id"].Value,
                        IsPartialAccent = isPartialAccent,
                        AccentRole = match.Groups["role"].Value.Trim().Trim('`')
                    });
                }
            }

            var commandIds = catalog
                // Hidden compatibility IDs still need valid assets and reviewer evidence.
                .Where(record => !record.IsObsolete &&
                    record.ExecutionKind != FeatureExecutionKind.Widget)
                .Select(record => record.Id)
                .OrderBy(id => id, StringComparer.Ordinal)
                .ToList();
            var reviewerIds = decisions
                .Select(decision => decision.FeatureId)
                .OrderBy(id => id, StringComparer.Ordinal)
                .ToList();
            var partialAccentIds = decisions
                .Where(decision => decision.IsPartialAccent)
                .Select(decision => decision.FeatureId)
                .OrderBy(id => id, StringComparer.Ordinal)
                .ToList();
            var neutralIds = decisions
                .Where(decision => !decision.IsPartialAccent)
                .Select(decision => decision.FeatureId)
                .OrderBy(id => id, StringComparer.Ordinal)
                .ToList();

            TestAssert.AreEqual(179, decisions.Count, "Reviewer batches must contain exactly 179 command decisions.");
            TestAssert.AreEqual(179, decisions.Select(decision => decision.FeatureId).Distinct(StringComparer.Ordinal).Count(),
                "Reviewer batches must contain exactly one decision per command ID.");
            TestAssert.AreEqual(88, partialAccentIds.Count, "Reviewer contract must contain 88 partial-accent command IDs.");
            TestAssert.AreEqual(91, neutralIds.Count, "Reviewer contract must contain 91 neutral command IDs.");
            CollectionAssert.AreEqual(commandIds, reviewerIds,
                "Reviewer decision IDs must exactly match the unique non-widget command catalog.");

            foreach (string widgetId in new[] { "BtnSizePanel", "BtnFontPanel", "BtnSearchBar" })
            {
                TestAssert.IsFalse(reviewerIds.Contains(widgetId), widgetId + " is a widget and must not enter the 179-command contract.");
                TestAssert.AreEqual(FeatureExecutionKind.Widget, catalog.Single(record => record.Id == widgetId).ExecutionKind,
                    widgetId + " must remain classified as a widget.");
            }

            var drawableNames = new HashSet<string>(
                new[] { "path", "rect", "circle", "ellipse", "line", "polyline", "polygon" },
                StringComparer.OrdinalIgnoreCase);
            var implementedAccentIds = new List<string>();
            string iconDirectory = Path.Combine(root, "oPenEfficiency", "UI", "Assets", "Icons");
            foreach (ReviewerIconAccentDecision decision in decisions)
            {
                string iconPath = Path.Combine(iconDirectory, decision.FeatureId + ".svg");
                TestAssert.IsTrue(File.Exists(iconPath), "Missing command icon: " + iconPath);
                XDocument document = XDocument.Load(iconPath, LoadOptions.None);
                var accentRoots = document
                    .Descendants()
                    .Where(IsExplicitAccentElement)
                    .ToList();

                if (!decision.IsPartialAccent)
                {
                    TestAssert.AreEqual(0, accentRoots.Count,
                        decision.FeatureId + " is reviewer-approved neutral-only but carries accent metadata.");
                    continue;
                }

                implementedAccentIds.Add(decision.FeatureId);
                TestAssert.IsTrue(accentRoots.Count > 0,
                    decision.FeatureId + " is reviewer-approved partial accent but has no accent metadata.");
                TestAssert.IsFalse(string.Equals(decision.AccentRole, "none", StringComparison.OrdinalIgnoreCase),
                    decision.FeatureId + " is partial accent but has no reviewer semantic role.");
                var expectedRoles = decision.AccentRole.Split(',')
                    .Select(value => value.Trim()).OrderBy(value => value, StringComparer.Ordinal).ToList();
                var actualRoles = new HashSet<string>(StringComparer.Ordinal);
                foreach (XElement accentRoot in accentRoots)
                {
                    XAttribute role = accentRoot.Attributes().FirstOrDefault(attribute =>
                        string.Equals(attribute.Name.LocalName, "data-oe-accent-role", StringComparison.OrdinalIgnoreCase));
                    TestAssert.IsNotNull(role,
                        decision.FeatureId + " has an accent element without an explicit semantic role.");
                    TestAssert.IsTrue(expectedRoles.Contains(role.Value.Trim()),
                        decision.FeatureId + " SVG role differs from the reviewer decision.");
                    actualRoles.Add(role.Value.Trim());
                }
                CollectionAssert.AreEqual(expectedRoles,
                    actualRoles.OrderBy(value => value, StringComparer.Ordinal).ToList(),
                    decision.FeatureId + " must implement exactly the reviewed semantic role set.");

                bool hasNeutralDrawable = document
                    .Descendants()
                    .Where(element => drawableNames.Contains(element.Name.LocalName))
                    .Any(element => !element.AncestorsAndSelf().Any(IsExplicitAccentElement));
                TestAssert.IsTrue(hasNeutralDrawable,
                    decision.FeatureId + " accents the whole icon; partial accents require neutral base geometry.");
            }

            implementedAccentIds.Sort(StringComparer.Ordinal);
            CollectionAssert.AreEqual(partialAccentIds, implementedAccentIds,
                "Implemented accent IDs must exactly match the reviewer-approved command IDs.");
        }

        private static bool IsExplicitAccentElement(XElement element)
        {
            XAttribute role = element.Attributes().FirstOrDefault(attribute =>
                string.Equals(attribute.Name.LocalName, "data-oe-role", StringComparison.OrdinalIgnoreCase));
            return role != null && string.Equals(role.Value.Trim(), "accent", StringComparison.OrdinalIgnoreCase);
        }

        private sealed class ReviewerIconAccentDecision
        {
            public string FeatureId { get; set; }
            public bool IsPartialAccent { get; set; }
            public string AccentRole { get; set; }
        }

        [DataContract]
        private sealed class IconAccentAuditRecord
        {
            [DataMember(Name = "featureId")]
            public string FeatureId { get; set; }

            [DataMember(Name = "scope")]
            public string Scope { get; set; }

            [DataMember(Name = "render16Light")]
            public string Render16Light { get; set; }

            [DataMember(Name = "render16Dark")]
            public string Render16Dark { get; set; }

            [DataMember(Name = "render24Light")]
            public string Render24Light { get; set; }

            [DataMember(Name = "render24Dark")]
            public string Render24Dark { get; set; }

            [DataMember(Name = "render48Light")]
            public string Render48Light { get; set; }

            [DataMember(Name = "render48Dark")]
            public string Render48Dark { get; set; }

            [DataMember(Name = "reviewerStatus")]
            public string ReviewerStatus { get; set; }
        }

        private static FeatureCatalogRecord ReadManualRecord(string source, string id, string name)
        {
            var escapedId = Regex.Escape(id);
            var match = Regex.Match(
                source,
                @"case\s+""" + escapedId +
                @""":\s*return\s+new\s+FeatureDisplayInfo\s*\{(.*?)\};",
                SourceOptions);
            if (!match.Success)
                throw new InvalidOperationException($"Manual catalog record '{id}' has no FeatureDisplayInfo metadata.");

            var body = match.Groups[1].Value;
            var record = new FeatureCatalogRecord
            {
                Id = id,
                Name = name,
                Tooltip = ReadString(body, "Tooltip"),
                Description = ReadString(body, "Description"),
                DetailedHelpText = ReadString(body, "DetailedHelpText"),
                SearchAliases = ReadStringArray(body, "SearchAliases"),
                MinSelection = ReadInt(body, "MinSelection"),
                MaxSelection = ReadInt(body, "MaxSelection"),
                RequiredType = ReadPowerPointSelectionType(body),
                RequiresTable = ReadBool(body, "RequiresTable"),
                RequiresHiddenShapes = ReadBool(body, "RequiresHiddenShapes"),
                IsToggle = ReadBool(body, "IsToggle"),
                UserVisible = true,
                SupportsUserShortcut = true,
                IsManual = true
            };
            FeatureCatalogPolicy.Apply(record);
            return record;
        }

        private static void ValidateCurrentPresets(
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            var definitions = SidebarPresetFactory.CreateDefinitions(catalog);

            AssertNoErrors(
                "generated presets",
                FeatureCatalogValidator.ValidatePresets(catalog, definitions));

            Assert(
                definitions.Select(definition => definition.Name).SequenceEqual(
                    new[]
                    {
                        SidebarPresetFactory.ConsultingStandardName,
                        SidebarPresetFactory.AllFeaturesName,
                        SidebarPresetFactory.MinimalName,
                        SidebarPresetFactory.LabsName
                    }),
                "The shipped preset names or order do not match the P0.3 specification.");

            foreach (var definition in definitions)
            {
                Assert(
                    definition.FeatureIds.Count ==
                    definition.FeatureIds.Distinct(StringComparer.Ordinal).Count(),
                    $"Preset '{definition.Name}' contains a duplicate command ID.");
            }

            var userVisibleIds = new HashSet<string>(
                catalog.Where(record => record.UserVisible && !record.IsObsolete)
                    .Select(record => record.Id),
                StringComparer.Ordinal);
            var allFeatures = definitions.Single(definition =>
                definition.Name == SidebarPresetFactory.AllFeaturesName);
            Assert(
                userVisibleIds.SetEquals(allFeatures.FeatureIds),
                "All Features does not contain every UserVisible catalog ID exactly once.");

            var consulting = definitions.Single(definition =>
                definition.Name == SidebarPresetFactory.ConsultingStandardName);
            Assert(
                consulting.FeatureIds.FirstOrDefault() == "BtnSearchBar",
                "Consulting Standard does not place Command Search first.");
            Assert(
                consulting.FeatureIds.All(id =>
                    catalog.Single(record => record.Id == id).Maturity == FeatureMaturity.Stable),
                "Consulting Standard includes a Preview or Experimental command.");
            Assert(
                !new HashSet<string>(consulting.FeatureIds, StringComparer.Ordinal)
                    .SetEquals(allFeatures.FeatureIds),
                "All Features is not observably different from Consulting Standard.");

            var consultingLayout = SidebarPresetFactory.Create(
                SidebarPresetFactory.ConsultingStandardName,
                catalog);
            var finalSection = consultingLayout.Sections.Last();
            Assert(
                finalSection.Name == "SETTINGS & HELP" &&
                finalSection.Features.Select(feature => feature.Id).SequenceEqual(
                    new[] { "BtnSettings", "BtnExploreFeatures" }),
                "Consulting Standard does not keep Settings and Feature Explorer in its final section.");

            var allFeaturesLayout = SidebarPresetFactory.Create(
                SidebarPresetFactory.AllFeaturesName,
                catalog);
            Assert(
                (!catalog.Any(record => record.UserVisible && record.Maturity == FeatureMaturity.Preview) ||
                 allFeaturesLayout.Sections.Any(section => section.Name.StartsWith("PREVIEW", StringComparison.Ordinal))) &&
                allFeaturesLayout.Sections.Any(section => section.Name.IndexOf("EXPERIMENTAL", StringComparison.Ordinal) >= 0),
                "All Features does not visibly separate Preview and Experimental commands.");

            var labs = definitions.Single(definition =>
                definition.Name == SidebarPresetFactory.LabsName);
            var expectedLabsIds = new HashSet<string>(
                catalog.Where(record =>
                        record.UserVisible &&
                        !record.IsObsolete &&
                        (record.Maturity == FeatureMaturity.Preview ||
                         record.Maturity == FeatureMaturity.Experimental))
                    .Select(record => record.Id),
                StringComparer.Ordinal);
            Assert(
                labs.FeatureIds.All(id =>
                {
                    var maturity = catalog.Single(record => record.Id == id).Maturity;
                    return maturity == FeatureMaturity.Preview ||
                        maturity == FeatureMaturity.Experimental;
                }),
                "Labs contains a command that is neither Preview nor Experimental.");
            Assert(
                expectedLabsIds.SetEquals(labs.FeatureIds),
                "Labs does not contain every Preview and Experimental command exactly once.");
            Assert(
                SidebarPresetFactory.GetDescription(SidebarPresetFactory.LabsName)
                    .IndexOf("may change", StringComparison.OrdinalIgnoreCase) >= 0,
                "Labs does not state that included behavior may change.");

            var minimal = definitions.Single(definition =>
                definition.Name == SidebarPresetFactory.MinimalName);
            Assert(
                SidebarPresetFactory.MinimalRequiredIds.All(minimal.FeatureIds.Contains),
                "Minimal is missing one or more required capabilities.");
        }

        private static void ValidateRuleCoverage()
        {
            var valid = ValidRecord("StableA");
            var duplicate = ValidRecord("StableA");
            var invalidSelection = ValidRecord("InvalidSelection");
            invalidSelection.MinSelection = 3;
            invalidSelection.MaxSelection = 2;
            var invalidCatalog = FeatureCatalogValidator.ValidateCatalog(
                new[] { valid, duplicate, invalidSelection });
            AssertContains(invalidCatalog, "Duplicate feature ID");
            AssertContains(invalidCatalog, "contradictory selection counts");
            AssertContains(invalidCatalog, "duplicates its Name in Tooltip");

            var experimental = ValidRecord("Experimental");
            experimental.Maturity = FeatureMaturity.Experimental;
            experimental.Category = FeatureCategory.Labs;
            experimental.DefaultPlacement = FeatureDefaultPlacement.Labs;
            var presets = new[]
            {
                new FeaturePresetDefinition
                {
                    Name = "Generated",
                    FeatureIds = new[] { "StableA", "StableA", "Missing" },
                    IsAutomaticallyGenerated = true
                },
                new FeaturePresetDefinition
                {
                    Name = "All Features",
                    FeatureIds = Array.Empty<string>(),
                    IsAllFeaturesPreset = true
                },
                new FeaturePresetDefinition
                {
                    Name = "Professional",
                    FeatureIds = new[] { "Experimental" },
                    IsProfessionalDefault = true
                },
                new FeaturePresetDefinition
                {
                    Name = "Labs",
                    FeatureIds = new[] { "StableA" },
                    IsLabsPreset = true,
                    RequiredFeatureIds = new[] { "MissingRequired" }
                }
            };
            var presetErrors = FeatureCatalogValidator.ValidatePresets(
                new[] { valid, experimental },
                presets);
            AssertContains(presetErrors, "unknown feature ID");
            AssertContains(presetErrors, "duplicate feature ID");
            AssertContains(presetErrors, "absent from All Features");
            AssertContains(presetErrors, "appears in professional default");
            AssertContains(presetErrors, "appears in Labs");
            AssertContains(presetErrors, "Required feature");

            var manualShadow = ValidRecord("StableA");
            manualShadow.IsManual = true;
            AssertContains(
                FeatureCatalogValidator.ValidateManualShadowing(
                    new[] { "StableA" },
                    new[] { manualShadow }),
                "without an override reason");

            var widget = ValidRecord("Widget");
            widget.ExecutionKind = FeatureExecutionKind.Widget;
            var invalidArguments = ValidRecord("InvalidArguments");
            var commandErrors = FeatureCatalogValidator.ValidateCommands(
                new[] { valid, widget, invalidArguments },
                new[]
                {
                    new FeatureCommandContract { Id = "StableA", IsToggle = true },
                    new FeatureCommandContract { Id = "StableA" },
                    new FeatureCommandContract { Id = "Widget" },
                    new FeatureCommandContract { Id = "Unknown" },
                    new FeatureCommandContract
                    {
                        Id = "InvalidArguments",
                        HasValidPrimaryArguments = false
                    }
                });
            AssertContains(commandErrors, "Duplicate command ID");
            AssertContains(commandErrors, "unknown feature ID");
            AssertContains(commandErrors, "must not be registered as a command");
            AssertContains(commandErrors, "toggle metadata");
            AssertContains(commandErrors, "invalid primary arguments");

            var missing = ValidRecord("MissingCommand");
            AssertContains(
                FeatureCatalogValidator.ValidateCommands(new[] { missing }, Array.Empty<FeatureCommandContract>()),
                "has no registered command");
        }

        private static void ValidateCommandCoverage(
            string root,
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            string generated = File.ReadAllText(Path.Combine(
                root,
                "oPenEfficiency",
                "UI",
                "Panels",
                "MainSidebar.GeneratedCommands.cs"));
            string specialized = File.ReadAllText(Path.Combine(
                root,
                "oPenEfficiency",
                "UI",
                "Panels",
                "MainSidebar.Commands.cs"));
            var actionIds = Regex.Matches(
                    generated + specialized,
                    @"Register(?:Action|HostAction)\(registry,\s*""(?<id>Btn[^""]+)""",
                    SourceOptions)
                .Cast<Match>()
                .Select(match => match.Groups["id"].Value)
                .Concat(new[] { "BtnRepeatShape", "BtnApplyTextTool" });
            var toggleIds = Regex.Matches(
                    specialized,
                    @"RegisterToggle\(registry,\s*""(?<id>Btn[^""]+)""",
                    SourceOptions)
                .Cast<Match>()
                .Select(match => match.Groups["id"].Value)
                .Concat(new[] { "BtnFlightMode" });
            var contracts = actionIds
                .Select(id => new FeatureCommandContract { Id = id })
                .Concat(toggleIds.Select(id => new FeatureCommandContract
                {
                    Id = id,
                    IsToggle = true
                }))
                .ToList();

            AssertNoErrors(
                "current command coverage",
                FeatureCatalogValidator.ValidateCommands(catalog, contracts));
            int expectedCommandCount = catalog.Count(record =>
                record.ExecutionKind != FeatureExecutionKind.Widget);
            Assert(contracts.Count == expectedCommandCount,
                $"Expected {expectedCommandCount} commands plus the catalog widgets, found {contracts.Count} commands.");
            Assert(contracts.Count(contract => contract.Id == "BtnMagicResizer" && contract.IsToggle) == 1,
                "Scale Selected Shapes is not registered exactly once as a toggle command.");
        }

        private static void ValidateCanonicalAliases(IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            var expected = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["Dock Left"] = "BtnDockLeft",
                ["Dock Right"] = "BtnDockRight",
                ["Dock Top"] = "BtnDockTop",
                ["Dock Bottom"] = "BtnDockBottom",
                ["Stack Top"] = "BtnDockTop",
                ["Stack Bottom"] = "BtnDockBottom",
                ["Presentation Style Check"] = "BtnStyleCheck",
                ["Style & Consistency Check"] = "BtnStyleCheck",
                ["Deep Presentation Cleaner"] = "BtnCleaner",
                ["Cleaner"] = "BtnCleaner",
                ["Status Stickers"] = "BtnIllustrativeSticker",
                ["Illustrative Sticker"] = "BtnIllustrativeSticker",
                ["Illustrative Stickers"] = "BtnIllustrativeSticker",
                ["Sync Across Slides"] = "BtnSyncObjects",
                ["Shape Synchronizer"] = "BtnSyncObjects",
                ["Agenda Wizard"] = "BtnAgendaWizard",
                ["Agenda Layout Creator"] = "BtnSaveAgendaLayout",
                ["Magic Color Erase"] = "BtnTransparentColor",
                ["Private oPE Comments"] = "BtnSlideNotes",
                ["Slide Notes"] = "BtnSlideNotes",
                ["Numeration"] = "BtnNumeration",
                ["Search Command Palette"] = "BtnSearchBar",
                ["Search features"] = "BtnSearchBar",
                ["Decr Horiz Spacing"] = "BtnDecreaseHorizontalSpacing",
                ["Decr Vert Spacing"] = "BtnDecreaseVerticalSpacing",
                ["Incr Horiz Spacing"] = "BtnIncreaseHorizontalSpacing",
                ["Incr Vert Spacing"] = "BtnIncreaseVerticalSpacing",
                ["Diagonal Align"] = "BtnDiagonalAlign",
                ["Arrange in Circle"] = "BtnArrangeCircle",
                ["Rotate"] = "BtnRotateShape",
                ["Shape Locking"] = "BtnShapeLock",
                ["Magic Resizer"] = "BtnMagicResizer",
                ["Scale Shapes & Spacing"] = "BtnMagicResizer",
                ["Scale Selection"] = "BtnMagicResizer",
                ["Bulk Format Replacer"] = "BtnReplaceFont",
                ["Format Shape Dialog"] = "BtnFormatShapeDialog",
                ["Insert Special Chars"] = "BtnSpecialChars",
                ["Number Formatter"] = "BtnFormatNumbers",
                ["Transparency"] = "BtnTransparency",
                ["Table Branding"] = "BtnTableBranding",
                ["Table Sorting"] = "BtnTableSort",
                ["Table Sum"] = "BtnTableSum",
                ["Split to Grid"] = "BtnSplitShape",
                ["HTML Slide Import"] = "BtnHtmlSlideImport",
                ["Lottery"] = "BtnWinnerPicker",
                ["Aspect Ratio Lock"] = "BtnLockDimensions",
                ["Object Connector"] = "BtnObjectConnector",
                ["Smart Corners"] = "BtnSmartCorners",
                ["Apply Table Heatmap Formatting"] = "BtnTableHeatmap",
                ["Add Sticker"] = "BtnAddSticker",
                ["Export Wizard"] = "BtnExportWizard",
                ["Map Wizard"] = "BtnMapWizard",
                ["Icon Browser"] = "BtnIconBrowser",
                ["Gantt Chart & Timeline Editor"] = "BtnGanttChart",
                ["Quick Font Panel"] = "BtnFontPanel",
                ["Hide / Show Master Objects"] = "BtnHideMasterObjects",
                ["openEfficiency Settings"] = "BtnSettings"
            };

            foreach (var alias in expected)
            {
                var resolved = FeatureCatalogSearch.ResolveByNameOrAlias(catalog, alias.Key);
                if (!string.Equals(resolved?.Id, alias.Value, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        $"Legacy alias '{alias.Key}' resolved to '{resolved?.Id ?? "<none>"}' instead of '{alias.Value}'.");
                }
            }

            var expectedNames = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["BtnDecreaseHorizontalSpacing"] = "Decrease Horizontal Spacing",
                ["BtnDecreaseVerticalSpacing"] = "Decrease Vertical Spacing",
                ["BtnIncreaseHorizontalSpacing"] = "Increase Horizontal Spacing",
                ["BtnIncreaseVerticalSpacing"] = "Increase Vertical Spacing",
                ["BtnDiagonalAlign"] = "Align Diagonally",
                ["BtnArrangeCircle"] = "Arrange Shapes in Circle",
                ["BtnRotateShape"] = "Rotate Shapes",
                ["BtnShapeLock"] = "Lock/Unlock Shapes",
                ["BtnMagicResizer"] = "Scale Selected Shapes",
                ["BtnReplaceFont"] = "Replace Formatting",
                ["BtnFormatShapeDialog"] = "Open Format Shape Pane",
                ["BtnSpecialChars"] = "Insert Special Characters",
                ["BtnFormatNumbers"] = "Format Numbers",
                ["BtnTransparency"] = "Adjust Transparency",
                ["BtnTableBranding"] = "Table Style Library",
                ["BtnTableSort"] = "Sort Table Cells",
                ["BtnTableSum"] = "Calculate Table Totals",
                ["BtnSplitShape"] = "Split Shape into Grid",
                ["BtnHtmlSlideImport"] = "Import HTML Slides",
                ["BtnWinnerPicker"] = "Random Winner Picker",
                ["BtnLockDimensions"] = "Lock/Unlock Aspect Ratio",
                ["BtnObjectConnector"] = "Connect Shapes",
                ["BtnSmartCorners"] = "Toggle Rounded Corners",
                ["BtnTableHeatmap"] = "Apply Table Heatmap",
                ["BtnAddSticker"] = "Add Corner Sticker",
                ["BtnExportWizard"] = "Export Presentation",
                ["BtnMapWizard"] = "Map Builder",
                ["BtnIconBrowser"] = "Icon Library",
                ["BtnGanttChart"] = "Gantt & Timeline Editor",
                ["BtnFontPanel"] = "Font Panel",
                ["BtnHideMasterObjects"] = "Hide/Show Master Objects",
                ["BtnSettings"] = "Settings"
            };

            foreach (var expectedName in expectedNames)
            {
                var record = catalog.SingleOrDefault(item => item.Id == expectedName.Key);
                Assert(record != null && record.Name == expectedName.Value,
                    $"Catalog name for '{expectedName.Key}' is '{record?.Name ?? "<none>"}' instead of '{expectedName.Value}'.");
            }
        }

        private static void ValidateNamingAndBehaviorConsistency(
            string root,
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            FeatureCatalogRecord adjustment = catalog.Single(record => record.Id == "BtnAlignShapeAdjustments");
            FeatureCatalogRecord legacyAdjustment = catalog.Single(record => record.Id == "BtnMatchAngles");
            Assert(adjustment.UserVisible && !legacyAdjustment.UserVisible,
                "Copy Shape Adjustments must be the sole visible adjustment-transfer command; BtnMatchAngles is compatibility-only.");
            Assert(adjustment.SearchAliases.Contains("Apply Anchor Shape Adjustments"),
                "The retired adjustment command name must remain an explicit search alias.");
            var resolvedLegacyAdjustment = FeatureCatalogSearch.ResolveByNameOrAlias(
                catalog.Where(record => record.UserVisible && !record.IsObsolete),
                "Apply Anchor Shape Adjustments");
            Assert(resolvedLegacyAdjustment?.Id == "BtnAlignShapeAdjustments",
                "The historical adjustment command name must resolve to Copy Shape Adjustments.");

            string[] expectedStackVariants = { "Stack Left", "Stack Right", "Stack Above", "Stack Below" };
            foreach (string id in new[] { "BtnDockLeft", "BtnDockRight", "BtnDockTop", "BtnDockBottom" })
            {
                Assert(catalog.Single(record => record.Id == id).Variants.SequenceEqual(expectedStackVariants),
                    id + " must expose only the canonical Stack variant vocabulary.");
            }

            FeatureCatalogRecord aspectRatio = catalog.Single(record => record.Id == "BtnLockDimensions");
            Assert(aspectRatio.Variants.SequenceEqual(new[]
                {
                    "Lock/Unlock Aspect Ratio (All Slides)",
                    "Lock Aspect Ratio (All Slides)",
                    "Unlock Aspect Ratio (All Slides)"
                }),
                "Aspect-ratio context actions must come from canonical catalog variants.");

            FeatureCatalogRecord thermometer = catalog.Single(record => record.Id == "BtnThermometer");
            Assert(thermometer.Name == "Thermometer" && thermometer.SearchAliases.Contains("Thermometer Chart"),
                "Thermometer must use the shape-object name while retaining its historical Chart alias.");

            string layoutMenu = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "MainSidebar.ContextMenus.Layout.cs"));
            Assert(!Regex.IsMatch(layoutMenu, @"GetCatalogVariantLabel\([^\r\n]+,\s*""Dock (Left|Right|Top|Bottom)""\)"),
                "Stack context-menu fallback text must not reintroduce Dock terminology.");

            string workflowMenu = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Panels", "MainSidebar.ContextMenus.WorkflowAndVisual.cs"));
            Assert(!workflowMenu.Contains("Toggle Aspect Ratio Lock (All Slides)"),
                "The legacy Aspect Ratio Lock phrase must not reappear as live context-menu copy.");

            string librarySource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "FeatureLibrary.cs"));
            var catalogById = catalog.ToDictionary(record => record.Id, StringComparer.Ordinal);
            foreach (Match seed in Regex.Matches(
                librarySource,
                @"new\s+SidebarFeature\s*\{\s*Id\s*=\s*""([^""]+)"",\s*Name\s*=\s*""([^""]+)""\s*\}",
                SourceOptions))
            {
                string id = seed.Groups[1].Value;
                string seedName = seed.Groups[2].Value;
                Assert(!catalogById.TryGetValue(id, out FeatureCatalogRecord record) || record.Name == seedName,
                    $"FeatureLibrary seed '{id}' duplicates stale display text '{seedName}' instead of canonical '{record?.Name}'.");
            }

            // Public source docs must cover the actual catalog, not private migration notes.
            string publicReference = File.ReadAllText(Path.Combine(root, "docs", "Features.md"));
            var documentedRows = Regex.Matches(publicReference,
                    @"^### (?<name>[^\r\n]+)\r?\n\r?\n\*\*Command ID:\*\* `(?<id>Btn[^`]+)`", RegexOptions.Multiline)
                .Cast<Match>().ToList();
            var documentedIds = documentedRows.Select(match => match.Groups["id"].Value).ToList();
            var visible = catalog.Where(record => record.UserVisible && !record.IsObsolete).ToList();
            TestAssert.AreEqual(documentedIds.Count, documentedIds.Distinct(StringComparer.Ordinal).Count(),
                "Public feature reference contains duplicate command IDs.");
            CollectionAssert.AreEquivalent(visible.Select(record => record.Id).ToList(), documentedIds,
                "Public feature reference must match every current visible command/widget ID.");
            foreach (var record in visible)
                TestAssert.AreEqual(record.Name,
                    documentedRows.Single(match => match.Groups["id"].Value == record.Id).Groups["name"].Value,
                    "Public feature reference has the wrong canonical name for ID " + record.Id);

        }

        private static void ValidateSavedLayoutNameMigration()
        {
            var catalog = new[]
            {
                new SidebarFeature { Id = "BtnCleaner", Name = "Presentation Cleaner" },
                new SidebarFeature { Id = "BtnDockTop", Name = "Stack Above" },
                new SidebarFeature { Id = "BtnAlignShapeAdjustments", Name = "Copy Shape Adjustments" },
                new SidebarFeature { Id = "BtnAlignToTableCell", Name = "Align to Cell" },
                new SidebarFeature { Id = "BtnStickyNoteManager", Name = "Sticky Note Manager" }
            };
            var config = new SidebarConfig
            {
                Sections = new List<SidebarSection>
                {
                    new SidebarSection
                    {
                        Name = "MY CUSTOM SECTION",
                        Features = new List<SidebarFeature>
                        {
                            new SidebarFeature { Id = "BtnCleaner", Name = "Deep Presentation Cleaner" },
                            new SidebarFeature { Id = "UnknownCustomCommand", Name = "Keep Me" },
                            new SidebarFeature { Id = "BtnDockTop", Name = "Stack Top" }
                        }
                    },
                    new SidebarSection
                    {
                        Name = "SECOND CUSTOM SECTION",
                        Features = new List<SidebarFeature>
                        {
                            new SidebarFeature { Id = "BtnAlignToTable", Name = "Old Align" },
                            new SidebarFeature { Id = "BtnManageStickyNotes", Name = "Old Sticky Note Manager" },
                            new SidebarFeature { Id = "BtnAlignShapeAdjustments", Name = "Copy Shape Adjustments" },
                            new SidebarFeature { Id = "BtnMatchAngles", Name = "Apply Anchor Shape Adjustments" },
                            new SidebarFeature { Id = "BtnSwapHorizontal", Name = "Obsolete" }
                        }
                    }
                }
            };

            var result = SidebarLayoutMigration.Migrate(config, catalog);
            Assert(result.Changed, "Saved-layout migration did not report its changes.");
            Assert(config.Sections.Count == 2, "Saved-layout migration changed the section count.");
            Assert(config.Sections[0].Name == "MY CUSTOM SECTION", "Saved-layout migration renamed a custom section.");
            Assert(config.Sections[1].Name == "SECOND CUSTOM SECTION", "Saved-layout migration renamed a custom section.");

            var first = config.Sections[0].Features;
            Assert(first.Count == 3, "Saved-layout migration changed the custom feature count.");
            Assert(first[0].Id == "BtnCleaner" && first[0].Name == "Presentation Cleaner",
                "Saved-layout migration did not refresh the cleaner name by stable ID.");
            Assert(first[1].Id == "UnknownCustomCommand" && first[1].Name == "Keep Me",
                "Saved-layout migration removed or changed an unknown custom command.");
            Assert(first[2].Id == "BtnDockTop" && first[2].Name == "Stack Above",
                "Saved-layout migration changed command ordering.");

            var second = config.Sections[1].Features;
            Assert(second.Count == 3, "Saved-layout migration did not remove only the obsolete command.");
            Assert(second[0].Id == "BtnAlignToTableCell" && second[0].Name == "Align to Cell",
                "Saved-layout migration did not preserve the legacy ID migration.");
            Assert(second[1].Id == "BtnStickyNoteManager" && second[1].Name == "Sticky Note Manager",
                "Saved-layout migration did not preserve the sticky-note manager ID migration.");
            Assert(second[2].Id == "BtnAlignShapeAdjustments" && second[2].Name == "Copy Shape Adjustments",
                "Saved-layout migration did not consolidate the retired adjustment command ID.");
            Assert(result.ConsolidatedLegacyCount == 1,
                "Saved-layout migration did not report the duplicate legacy adjustment entry it consolidated.");
            Assert(result.UnknownFeatureIds.SequenceEqual(new[] { "UnknownCustomCommand" }),
                "Saved-layout migration did not report unknown IDs.");
        }

        private static void ValidatePresetGenerationPreservesCustomizedLayout(
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            var customized = new SidebarConfig
            {
                Sections = new List<SidebarSection>
                {
                    new SidebarSection
                    {
                        Name = "MY DAILY FLOW",
                        Color = "#123456",
                        IsVisible = false,
                        Features = new List<SidebarFeature>
                        {
                            new SidebarFeature { Id = "BtnMatchSize", Name = "Match Size" },
                            new SidebarFeature { Id = "BtnAlignLeft", Name = "Align Left" }
                        }
                    }
                }
            };
            var originalSection = customized.Sections[0];
            var originalFeatures = originalSection.Features.ToList();

            foreach (var presetName in SidebarPresetFactory.PresetNames)
                SidebarPresetFactory.Create(presetName, catalog);

            var resolved = SidebarPresetFactory.ResolveInitialLayout(customized, catalog);
            Assert(ReferenceEquals(resolved, customized),
                "Loading preset definitions replaced an existing customized layout.");
            Assert(customized.Sections.Count == 1, "Generating presets replaced a customized layout.");
            Assert(ReferenceEquals(customized.Sections[0], originalSection),
                "Generating presets replaced a customized section.");
            Assert(customized.Sections[0].Name == "MY DAILY FLOW" &&
                customized.Sections[0].Color == "#123456" &&
                !customized.Sections[0].IsVisible,
                "Generating presets changed customized section properties.");
            Assert(customized.Sections[0].Features.SequenceEqual(originalFeatures),
                "Generating presets changed customized feature ordering.");

            var emptyCustomized = new SidebarConfig { Sections = new List<SidebarSection>() };
            Assert(
                ReferenceEquals(
                    SidebarPresetFactory.ResolveInitialLayout(emptyCustomized, catalog),
                    emptyCustomized),
                "Loading preset definitions replaced an intentionally empty customized layout.");

            var cleanInstall = SidebarPresetFactory.ResolveInitialLayout(null, catalog);
            Assert(
                cleanInstall.Sections.SelectMany(section => section.Features)
                    .First().Id == "BtnSearchBar",
                "A clean installation does not receive Consulting Standard.");
        }

        private static FeatureCatalogRecord ValidRecord(string id)
        {
            return new FeatureCatalogRecord
            {
                Id = id,
                Name = id,
                Tooltip = id,
                Description = $"{id} description.",
                DetailedHelpText = $"{id} help.",
                Category = FeatureCategory.AlignAndSize,
                Maturity = FeatureMaturity.Stable,
                UserVisible = true,
                DefaultPlacement = FeatureDefaultPlacement.CatalogOnly,
                SupportsUserShortcut = true,
                RequiredType = OfficeSelectionType.Shapes
            };
        }

        private static string ReadString(string source, string property)
        {
            var match = Regex.Match(
                source,
                @"(?<![A-Za-z0-9_])" + Regex.Escape(property) + @"\s*=\s*""((?:\\.|[^""])*)""",
                SourceOptions);
            return match.Success ? Regex.Unescape(match.Groups[1].Value) : null;
        }

        private static IReadOnlyList<string> ReadStringArray(string source, string property)
        {
            var match = Regex.Match(
                source,
                @"(?<![A-Za-z0-9_])" + Regex.Escape(property) + @"\s*=\s*new(?:\s+string)?\[\]\s*\{(.*?)\}",
                SourceOptions);
            if (!match.Success) return Array.Empty<string>();

            return Regex.Matches(match.Groups[1].Value, @"""((?:\\.|[^""])*)""")
                .Cast<Match>()
                .Select(item => Regex.Unescape(item.Groups[1].Value))
                .ToList();
        }

        private static int ReadInt(string source, string property)
        {
            var match = Regex.Match(
                source,
                Regex.Escape(property) + @"\s*=\s*(-?\d+)",
                SourceOptions);
            return match.Success ? int.Parse(match.Groups[1].Value) : 0;
        }

        private static bool ReadBool(string source, string property)
        {
            return Regex.IsMatch(
                source,
                Regex.Escape(property) + @"\s*=\s*true",
                RegexOptions.CultureInvariant);
        }

        private static bool HasExplicitFalse(string source, string property)
        {
            return Regex.IsMatch(
                source,
                Regex.Escape(property) + @"\s*=\s*false",
                RegexOptions.CultureInvariant);
        }

        private static OfficeSelectionType ReadSelectionType(string source)
        {
            var match = Regex.Match(
                source,
                @"RequiredType\s*=\s*OfficeSelectionType\.(\w+)",
                RegexOptions.CultureInvariant);
            return match.Success
                ? (OfficeSelectionType)Enum.Parse(typeof(OfficeSelectionType), match.Groups[1].Value)
                : OfficeSelectionType.Shapes;
        }

        private static OfficeSelectionType ReadPowerPointSelectionType(string source)
        {
            var match = Regex.Match(
                source,
                @"RequiredType\s*=\s*PowerPoint\.PpSelectionType\.(\w+)",
                RegexOptions.CultureInvariant);
            if (!match.Success) return OfficeSelectionType.None;

            switch (match.Groups[1].Value)
            {
                case "ppSelectionShapes": return OfficeSelectionType.Shapes;
                case "ppSelectionText": return OfficeSelectionType.Text;
                case "ppSelectionSlides": return OfficeSelectionType.Slides;
                default: return OfficeSelectionType.None;
            }
        }

        private static void AssertNoErrors(string scope, IReadOnlyList<string> errors)
        {
            if (errors.Count == 0) return;
            throw new InvalidOperationException(
                $"{scope} validation failed:{Environment.NewLine}- " +
                string.Join(Environment.NewLine + "- ", errors));
        }

        private static void AssertContains(IEnumerable<string> errors, string expected)
        {
            if (errors.Any(error => error.IndexOf(expected, StringComparison.OrdinalIgnoreCase) >= 0))
                return;
            throw new InvalidOperationException(
                $"Validator rule coverage did not produce an error containing '{expected}'.");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        private static void ValidateWpfDispatcherFailureFallback()
        {
            string root = FindRepositoryRoot();
            string addIn = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "ThisAddIn.cs"));
            string registry = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Services", "Commands", "CommandRegistry.cs"));

            Assert(addIn.Contains("app.DispatcherUnhandledException += WpfApplication_DispatcherUnhandledException") &&
                   addIn.Contains("_wpfDispatcherHandlerRegistered = true"),
                "The WPF dispatcher fallback is not registered exactly once for non-registry UI entry points.");
            Assert(Regex.IsMatch(
                    addIn,
                    @"WpfApplication_DispatcherUnhandledException\([\s\S]*?e\.Handled = true;[\s\S]*?ExceptionLogger\.Log\(e\.Exception, \""ThisAddIn\.WpfDispatcherUnhandledException\""\)[\s\S]*?CommandRegistry\.UserVisibleFailureMessage",
                    RegexOptions.CultureInvariant),
                "Unhandled WPF UI failures are not marked handled, logged, and surfaced through the shared command failure message.");
            Assert(registry.Contains("internal const string UserVisibleFailureMessage"),
                "The command failure message is not shared with the WPF dispatcher fallback.");
        }

        private static string ExtractMethodBody(string source, string signature, string methodName)
        {
            int signatureStart = source.IndexOf(signature, StringComparison.Ordinal);
            Assert(signatureStart >= 0, methodName + " signature could not be located.");
            Assert(source.IndexOf(signature, signatureStart + signature.Length, StringComparison.Ordinal) < 0,
                methodName + " signature is ambiguous; update the source contract.");

            int bodyStart = source.IndexOf('{', signatureStart + signature.Length);
            Assert(bodyStart >= 0, methodName + " body start could not be located.");

            int depth = 0;
            for (int index = bodyStart; index < source.Length; index++)
            {
                if (source[index] == '{') depth++;
                if (source[index] != '}') continue;

                depth--;
                if (depth == 0) return source.Substring(bodyStart, index - bodyStart + 1);
            }

            throw new InvalidOperationException(methodName + " body end could not be located.");
        }

        private static void ValidateSelectionMetadataWindowProbe()
        {
            string root = FindRepositoryRoot();
            string source = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Services", "PowerPointManager.cs"));
            string body = ExtractMethodBody(
                source,
                "public SelectionMetadata GetCurrentSelectionMetadata()",
                "PowerPointManager.GetCurrentSelectionMetadata");

            Assert(body.Contains("ExceptionPolicy.TryComProbe("),
                "Selection metadata must treat a missing active PowerPoint window as an expected COM probe failure.");
            Assert(body.Contains("out DocumentWindow window") && body.Contains("var selection = window.Selection;"),
                "Selection metadata must reuse the successfully probed PowerPoint window.");
            Assert(!body.Contains("_application.ActiveWindow == null") &&
                   Regex.Matches(body, @"_application\.ActiveWindow").Count == 1,
                "Selection metadata must not repeat the ActiveWindow COM boundary outside the guarded probe.");
            Assert(body.Contains("ExceptionPolicy.ReportOptionalComFailure(ex, \"PowerPointManager.GetCurrentSelectionMetadata\")") &&
                   !body.Contains("ExceptionLogger.Log(ex, \"PowerPointManager.GetCurrentSelectionMetadata\")"),
                "Selection metadata must classify expected COM failures on the probed window as optional and log only unexpected exceptions.");
        }

        private static void ValidateSidebarPaneCleanupDisposalBoundary()
        {
            string root = FindRepositoryRoot();
            string source = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Services", "SidebarPaneCoordinator.cs"));
            string body = ExtractMethodBody(
                source,
                "private void RemovePane(CustomTaskPane pane, string context)",
                "SidebarPaneCoordinator.RemovePane");

            Assert(body.Contains("ExceptionPolicy.RunBestEffort("),
                "Sidebar pane removal is best-effort cleanup and must run through the exception policy.");
            Assert(body.Contains("ObjectDisposedException") &&
                   body.Contains("ExceptionPolicy.IsExpectedComFailure("),
                "Sidebar pane removal must treat a disposed CustomTaskPanes collection as an expected cleanup miss.");
            Assert(!body.Contains("catch (COMException)") && !body.Contains("catch (Exception ex) { ExceptionLogger.Log"),
                "Sidebar pane removal must not bypass the policy with ad-hoc catches.");
        }

        private static void ValidateAlignAnchorOracleConfiguration()
        {
            string root = FindRepositoryRoot();
            string managerSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "Services", "PowerPointManager.cs"));
            string seamSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "ThisAddIn.Automation.cs"));
            string anchorModesSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency.Core", "Utils", "AnchorModes.cs"));
            string applyTextSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "Dialogs", "ApplyTextWindow.xaml.cs"));
            string settingsSource = File.ReadAllText(Path.Combine(
                root, "oPenEfficiency", "UI", "SettingsViewModel.cs"));
            string referenceBody = ExtractMethodBody(
                managerSource,
                "public Shape GetReferenceShape()",
                "PowerPointManager.GetReferenceShape");
            string anchorModeBody = ExtractMethodBody(
                managerSource,
                "public bool UsesLastSelectedAnchor()",
                "PowerPointManager.UsesLastSelectedAnchor");
            string anchorModesBody = ExtractMethodBody(
                anchorModesSource,
                "public static bool IsLastSelected(string masterObjectMode)",
                "AnchorModes.IsLastSelected");
            string seamBody = ExtractMethodBody(
                seamSource,
                "internal string ExerciseAlignToFirstForAutomation(string direction, int shapeCount)",
                "ThisAddIn.ExerciseAlignToFirstForAutomation");

            Assert(referenceBody.Contains("UsesLastSelectedAnchor()"),
                "The reference-shape rule must read its first/last decision from the single anchor-mode decision point.");
            Assert(anchorModeBody.Contains("AnchorModes.IsLastSelected(") &&
                   anchorModeBody.Contains("MasterObjectMode"),
                "The anchor-mode decision must delegate the string mapping to the shared host-neutral predicate.");
            Assert(anchorModesBody.Contains("\"Last Selected\"") &&
                   anchorModesBody.Contains("StringComparison.OrdinalIgnoreCase"),
                "The shared anchor-mode predicate must pin the exact configured string mapping ('Last Selected', OrdinalIgnoreCase).");
            Assert(seamBody.Contains("JsonService.Load") &&
                   seamBody.Contains("\"Last Selected\"") &&
                   seamBody.Contains("StringComparison.OrdinalIgnoreCase") &&
                   seamBody.Contains("anchorMatchesConfiguration =") &&
                   seamBody.Contains("anchorMatchesConfiguration &&"),
                "The align-first seam must derive its anchor expectation from the raw configured mode, not from the production resolver.");
            Assert(seamBody.Contains("anchorModeAgreesWithProduction =") &&
                   seamBody.Contains("anchorModeAgreesWithProduction &&"),
                "The align-first seam must cross-check the raw configured mode against the production anchor-mode decision.");
            Assert(applyTextSource.Contains("_manager.UsesLastSelectedAnchor()") &&
                   !applyTextSource.Contains("\"Last Selected\""),
                "ApplyTextWindow must route its first/last decision through the production decision point instead of re-deciding the configured mode locally.");
            Assert(settingsSource.Contains("AnchorModes.IsLastSelected(Config.MasterObjectMode)") &&
                   !settingsSource.Contains("\"Last Selected\""),
                "SettingsViewModel must canonicalize the configured mode through the shared anchor-mode predicate.");
        }

        private static void ValidateBuildStampIdentityComparison()
        {
            string root = FindRepositoryRoot();
            string gateSource = File.ReadAllText(Path.Combine(
                root, "eng", "Test-VstoDeploymentIdentity.ps1"));

            Assert(gateSource.Contains("[System.Version]::Parse($BuildStamp)"),
                "The deployment identity gate must compare parsed Version identities, not raw stamp strings.");
            Assert(gateSource.Contains("$assemblyVersion = [System.Reflection.AssemblyName]::GetAssemblyName($assemblyPath).Version") &&
                   gateSource.Contains("if ($assemblyVersion -ne $expectedAssemblyVersion)"),
                "The deployment identity gate must compare the read-back assembly Version against the parsed stamp.");
            Assert(!gateSource.Contains("-cne $BuildStamp"),
                "The deployment identity gate must not compare the read-back assembly version against the raw stamp string (zero-padded stamp parts normalize away and false-fail builds between 00:00 and 09:59).");
            Assert(System.Version.Parse("2026.9.28.0020") == System.Version.Parse("2026.9.28.20"),
                ".NET Version parsing must normalize zero-padded stamp components identically on both sides of the identity comparison.");
        }

        private static void ValidateInstallerIdentityModeHandOff()
        {
            string root = FindRepositoryRoot();
            string buildSource = File.ReadAllText(Path.Combine(root, "eng", "Build.ps1"));
            string propsSource = File.ReadAllText(Path.Combine(root, "Directory.Build.props"));
            string targetsSource = File.ReadAllText(Path.Combine(root, "Directory.Build.targets"));

            Assert(buildSource.Contains("$env:OPENEFFICIENCY_USE_COMMITTED_VSTO_APPLICATION_VERSION = $useCommittedVstoApplicationVersion"),
                "The canonical build must hand its VSTO identity mode (per-build stamp vs committed release CalVer) to the nested builds it starts.");
            Assert(buildSource.Contains("Remove-Item Env:\\OPENEFFICIENCY_USE_COMMITTED_VSTO_APPLICATION_VERSION"),
                "The canonical build must remove the identity-mode hand-off again in its finally block so a stray child build keeps the documented per-build fallback.");
            Assert(propsSource.Contains("OPENEFFICIENCY_USE_COMMITTED_VSTO_APPLICATION_VERSION") &&
                   propsSource.Contains("<UseCommittedVstoApplicationVersion Condition=\"'$(UseCommittedVstoApplicationVersion)' == ''"),
                "Directory.Build.props must map the identity-mode hand-off into UseCommittedVstoApplicationVersion when the property is not set explicitly.");
            Assert(targetsSource.Contains("'$(UseCommittedVstoApplicationVersion)' != 'true'"),
                "Directory.Build.targets must only replace ApplicationVersion with the per-build stamp while the installer identity mode is off.");
            Assert(buildSource.Contains("-UseCommittedApplicationVersion:$Installer"),
                "The deployment identity gate must validate the same identity mode the build stamped.");
        }

        private static void ValidateIllustrationEditableInsertionNaming()
        {
            string root = FindRepositoryRoot();
            string service = File.ReadAllText(Path.Combine(root, "oPenEfficiency", "Services", "IllustrationInsertionService.cs"));

            Assert(service.Contains("shape.Name = BuildShapeName(item);"),
                "The inserted picture must carry the illustration's source name (Selection Pane provenance).");
            Assert(service.Contains("converted.Name = BuildShapeName(item);"),
                "The editable-shapes route must re-name the converted result with the same source name, so the Selection Pane keeps showing the illustration's provenance after SVGEdit replaces the picture (owner finding 2026-10-19, REG-ILLUSTRATION-EDITNAME-001).");
            Assert(service.Contains("MatchOriginalBounds(converted") && service.Contains("TagShape(converted, item);"),
                "The editable-shapes route must re-fit and re-tag the converted result inside the cosmetics undo unit.");
            Assert(service.Contains("return \"Illustration \" + name;") &&
                   service.Contains("string name = string.IsNullOrWhiteSpace(item.Name) ? item.Id : item.Name;"),
                "Both insertion routes must share the single source-name builder (item name with id fallback).");
            Assert(service.Contains("ShapeMetadataService.SetTag(shape, \"oPE_Illustration\", item.Id);"),
                "The oPE_Illustration provenance is a shape TAG holding the item id - it is never a shape name and never shows in the Selection Pane.");
        }

        private static void ValidateAcceptanceRunnerArgumentBinding()
        {
            string root = FindRepositoryRoot();
            string driver = File.ReadAllText(Path.Combine(root, "eng", "Invoke-FeatureAcceptance.ps1"));

            Assert(!driver.Contains("& $scriptPath @arguments"),
                "Acceptance runner arguments must never be splatted as an array: array splatting passes elements positionally, so a token like '-Commands' arrives as a literal value and the runner's parameter binder rejects it (REG-ACCEPTANCE-ARGS-001).");
            Assert(driver.Contains("$namedArguments[$parameterName]") && driver.Contains("& $scriptPath @namedArguments"),
                "Acceptance runner arguments must be parsed into named parameters and splatted as a hashtable so manifest tokens bind to the runner's real parameters.");
            Assert(driver.Contains("cannot be bound safely"),
                "A runner argument token that is not a parameter name must fail loudly instead of binding positionally and under-testing the scenario.");
            Assert(driver.Contains("$namedArguments[$parameterName] = $true"),
                "A parameter token without a value must bind as a switch (e.g. -RequireAddIn), and a token with several values must bind as the value list.");
        }

        private static string FindRepositoryRoot()
        {
            var current = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (current != null)
            {
                if (File.Exists(Path.Combine(current.FullName, "oPenEfficiency.slnx")))
                    return current.FullName;
                current = current.Parent;
            }

            throw new DirectoryNotFoundException("Could not locate the repository root.");
        }
    }
}
