using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Excel.Services;
using oPenEfficiency.Services;
using oPenEfficiency.Features.Utilities;

namespace oPenEfficiency.Excel.Tests
{
    /// <summary>
    /// FEAT-005 S0 / DL-S0 compatibility contract tests. They freeze the legacy
    /// link payload readers and the PowerPoint multi-instance policy exactly as
    /// documented in docs/product/feat-005-link-format-compatibility-contract.md.
    /// </summary>
    [TestClass]
    public class LinkCompatibilityContractTests
    {
        // --- TableLink payload (contract §1.2) ---

        [TestMethod] public void TableLinkFlagsRoundTripTest() => TableLinkFlagsRoundTrip();

        private static void TableLinkFlagsRoundTrip()
        {
            var cases = new[]
            {
                new { Formatted = true,  AutoResize = false, Flags = "F" },
                new { Formatted = false, AutoResize = true,  Flags = "R" },
                new { Formatted = true,  AutoResize = true,  Flags = "FR" },
                new { Formatted = false, AutoResize = false, Flags = "V" }
            };
            foreach (var item in cases)
            {
                var link = new TableLink
                {
                    Path = @"C:\Decks\FY26.xlsx",
                    Sheet = "P&L",
                    Range = "A1:C4",
                    Formatted = item.Formatted,
                    AutoResize = item.AutoResize
                };
                string payload = link.Build();
                Assert.IsTrue(payload.StartsWith("oPE_TableLink|" + item.Flags + "|", StringComparison.Ordinal),
                    "Flags token must serialize as " + item.Flags + " for payload: " + payload);
                Assert.IsTrue(TableLink.TryParse(payload, out var parsed), "Built payload must parse: " + payload);
                Assert.AreEqual(item.Formatted, parsed.Formatted, "Formatted flag must round-trip.");
                Assert.AreEqual(item.AutoResize, parsed.AutoResize, "AutoResize flag must round-trip.");
                Assert.AreEqual(link.Path, parsed.Path, "Path must round-trip.");
                Assert.AreEqual(link.Sheet, parsed.Sheet, "Sheet must round-trip.");
                Assert.AreEqual(link.Range, parsed.Range, "Range must round-trip.");
            }
        }

        [TestMethod] public void TableLinkLegacyFlaglessPayloadParsesTest() => TableLinkLegacyFlaglessPayloadParses();

        private static void TableLinkLegacyFlaglessPayloadParses()
        {
            Assert.IsTrue(TableLink.TryParse(
                "oPE_TableLink|C:\\Data\\Book1.xlsx|Summary|B2:D9", out var legacy),
                "Legacy flag-less payload must keep parsing.");
            Assert.AreEqual(@"C:\Data\Book1.xlsx", legacy.Path);
            Assert.AreEqual("Summary", legacy.Sheet);
            Assert.AreEqual("B2:D9", legacy.Range);
            Assert.IsFalse(legacy.Formatted, "Legacy payload defaults to values-only.");
            Assert.IsFalse(legacy.AutoResize, "Legacy payload defaults to fixed size.");
        }

        [TestMethod] public void TableLinkSheetNameWithPipesRejoinsTest() => TableLinkSheetNameWithPipesRejoins();

        private static void TableLinkSheetNameWithPipesRejoins()
        {
            var link = new TableLink
            {
                Path = @"C:\Data\Book1.xlsx",
                Sheet = "Q1|Q2|Consolidated",
                Range = "A1:B2"
            };
            Assert.IsTrue(TableLink.TryParse(link.Build(), out var parsed));
            Assert.AreEqual("Q1|Q2|Consolidated", parsed.Sheet,
                "Pipes inside the sheet name must be rejoined.");
        }

        // --- ExcelValueLink payload (contract §1.3) ---

        [TestMethod] public void ExcelValueLinkRoundTripTest() => ExcelValueLinkRoundTrip();

        private static void ExcelValueLinkRoundTrip()
        {
            var link = new ExcelValueLink
            {
                Path = @"C:\Decks\FY26 model.xlsx",
                Sheet = "Assumptions",
                Cell = "$B$3",
                Decimals = 1,
                Thousands = true,
                Percent = false,
                Raw = 1234.56,
                Id = "abc123",
                Start = 12,
                Length = 6,
                Text = "1,234.6"
            };
            Assert.IsTrue(ExcelValueLink.TryParse(link.Build(), out var parsed));
            Assert.AreEqual(link.Path, parsed.Path);
            Assert.AreEqual(link.Sheet, parsed.Sheet);
            Assert.AreEqual(link.Cell, parsed.Cell);
            Assert.AreEqual(link.Decimals, parsed.Decimals);
            Assert.IsTrue(parsed.Thousands);
            Assert.IsFalse(parsed.Percent);
            Assert.AreEqual(link.Raw, parsed.Raw);
            Assert.AreEqual(link.Id, parsed.Id);
            Assert.AreEqual(link.Start, parsed.Start);
            Assert.AreEqual(link.Length, parsed.Length);
            Assert.AreEqual(link.Text, parsed.Text);
        }

        [TestMethod] public void ExcelValueLinkLegacyMinimalPayloadParsesTest() => ExcelValueLinkLegacyMinimalPayloadParses();

        private static void ExcelValueLinkLegacyMinimalPayloadParses()
        {
            Assert.IsTrue(ExcelValueLink.TryParse(
                "oPE_ExcelValue\npath=C:\\Data\\Book1.xlsx\nsheet=Sheet1\ncell=$A$1",
                out var legacy),
                "Pre-id/start/len payloads must keep parsing.");
            Assert.AreEqual(@"C:\Data\Book1.xlsx", legacy.Path);
            Assert.AreEqual("Sheet1", legacy.Sheet);
            Assert.AreEqual("$A$1", legacy.Cell);
            Assert.AreEqual(-1, legacy.Decimals, "Missing dec keeps the as-displayed default.");
            Assert.IsNull(legacy.Id, "Missing id stays null.");
            Assert.AreEqual(-1, legacy.Start, "Missing start stays -1 (not inline).");
            Assert.IsNull(legacy.Raw, "Missing raw stays null.");
        }

        [TestMethod] public void ExcelValueLinkTextWithEqualsParsesTest() => ExcelValueLinkTextWithEqualsParses();

        private static void ExcelValueLinkTextWithEqualsParses()
        {
            var link = new ExcelValueLink { Path = "p", Sheet = "s", Cell = "$A$1", Text = "a=b=c" };
            Assert.IsTrue(ExcelValueLink.TryParse(link.Build(), out var parsed));
            Assert.AreEqual("a=b=c", parsed.Text,
                "Only the first '=' separates key from value; the text line may contain '='.");
        }

        [TestMethod] public void ExcelValueLinkFormatValueModesTest() => ExcelValueLinkFormatValueModes();

        private static void ExcelValueLinkFormatValueModes()
        {
            // Contract fact: FormatValue renders with the CURRENT culture
            // (ExcelValueFeature uses CultureInfo.CurrentCulture), not invariant.
            var previous = System.Threading.Thread.CurrentThread.CurrentCulture;
            try
            {
                System.Threading.Thread.CurrentThread.CurrentCulture =
                    System.Globalization.CultureInfo.GetCultureInfo("en-US");

                var asDisplayed = new ExcelValueLink { Decimals = -1, Text = "42 %" };
                Assert.AreEqual("42 %", asDisplayed.FormatValue(),
                    "dec=-1 must show the stored Excel display text as-is.");

                var decimalValue = new ExcelValueLink { Decimals = 2, Thousands = true, Raw = 1234.5 };
                Assert.AreEqual("1,234.50", decimalValue.FormatValue());

                var percent = new ExcelValueLink { Decimals = 0, Percent = true, Raw = 0.456 };
                Assert.AreEqual("46 %", percent.FormatValue(),
                    "Percent mode scales the raw fraction by 100.");
            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = previous;
            }
        }

        // --- Fill parity policy (contract §2, DL-FORMAT-NOFILL-001) ---

        [TestMethod] public void NoFillPolicyIsExplicitWhiteTest() => NoFillPolicyIsExplicitWhite();

        private static void NoFillPolicyIsExplicitWhite()
        {
            Assert.AreEqual(0xFFFFFF, DataLinkFillPolicy.NoFillBgr,
                "Excel 'No Fill' resolves to explicit white; an invisible fill would expose the table style.");

            Assert.AreEqual(DataLinkFillPolicy.NoFillBgr, DataLinkFillPolicy.ResolveFillBgr(false, 0x123456),
                "A cell without an Excel fill must ignore the color passed alongside it.");

            Assert.AreEqual(0x0000FF, DataLinkFillPolicy.ResolveFillBgr(true, 0x0000FF),
                "An explicit Excel fill keeps its own BGR color (0x0000FF = red).");
        }

        // --- PowerPoint multi-instance policy (contract §3) ---

        [TestMethod] public void InstancePolicyFailsClosedOnMultipleProcessesTest() => InstancePolicyFailsClosedOnMultipleProcesses();

        private static void InstancePolicyFailsClosedOnMultipleProcesses()
        {
            Assert.AreEqual(PowerPointInstanceSelection.Ambiguous,
                PowerPointInstancePolicy.SelectProcessIds(new uint[] { 100, 200 }),
                "The process-level guard must reject two POWERPNT processes before COM attachment.");

            Assert.AreEqual(PowerPointInstanceSelection.None,
                PowerPointInstancePolicy.Select(new PowerPointWindowCandidate[0]),
                "No running PowerPoint must yield None.");

            Assert.AreEqual(PowerPointInstanceSelection.Single,
                PowerPointInstancePolicy.Select(new[]
                {
                    new PowerPointWindowCandidate(new IntPtr(1), 100),
                    new PowerPointWindowCandidate(new IntPtr(2), 100)
                }),
                "Several windows of the SAME process must yield Single.");

            Assert.AreEqual(PowerPointInstanceSelection.Ambiguous,
                PowerPointInstancePolicy.Select(new[]
                {
                    new PowerPointWindowCandidate(new IntPtr(1), 100),
                    new PowerPointWindowCandidate(new IntPtr(2), 200)
                }),
                "Two distinct PowerPoint processes must fail closed as Ambiguous.");
        }

        // --- Connector attach routing (contract \u00a73, DL-CONNECTOR-ROT-001) ---

        [TestMethod] public void SingleProcessWithoutVisibleFrameFallsBackToRotTest() => SingleProcessWithoutVisibleFrameFallsBackToRot();

        private static void SingleProcessWithoutVisibleFrameFallsBackToRot()
        {
            Assert.AreEqual(PowerPointAttachRoute.RunningObjectTableFallback,
                PowerPointInstancePolicy.SelectRoute(PowerPointInstanceSelection.Single, PowerPointInstanceSelection.None),
                "One live process with no enumerable visible frame must route to the PID-checked ROT fallback, not to a silent null.");

            Assert.AreEqual(PowerPointAttachRoute.VisibleWindow,
                PowerPointInstancePolicy.SelectRoute(PowerPointInstanceSelection.Single, PowerPointInstanceSelection.Single),
                "One process with a visible frame attaches through the frame window.");

            Assert.AreEqual(PowerPointAttachRoute.Ambiguous,
                PowerPointInstancePolicy.SelectRoute(PowerPointInstanceSelection.Ambiguous, PowerPointInstanceSelection.None),
                "Two distinct live processes must still fail closed.");

            Assert.AreEqual(PowerPointAttachRoute.Ambiguous,
                PowerPointInstancePolicy.SelectRoute(PowerPointInstanceSelection.Single, PowerPointInstanceSelection.Ambiguous),
                "Visible frames of two different processes must fail closed before any attach attempt.");

            Assert.AreEqual(PowerPointAttachRoute.None,
                PowerPointInstancePolicy.SelectRoute(PowerPointInstanceSelection.None, PowerPointInstanceSelection.None),
                "No running PowerPoint stays None.");
        }
    }
}
