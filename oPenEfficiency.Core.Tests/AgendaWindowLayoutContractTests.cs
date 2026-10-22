using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class AgendaWindowLayoutContractTests
    {
        private static readonly XNamespace Wpf = "http://schemas.microsoft.com/winfx/2006/xaml/presentation";
        private static readonly XNamespace Xaml = "http://schemas.microsoft.com/winfx/2006/xaml";
        private static readonly string[] AgendaFiles =
        {
            "oPenEfficiency/UI/Wizards/AgendaWizard.xaml",
            "oPenEfficiency/UI/Dialogs/SaveAgendaLayoutWindow.xaml",
            "oPenEfficiency/UI/ImportAgendaLayoutWindow.xaml"
        };

        private static string Root()
        {
            var directory = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (directory != null && !File.Exists(Path.Combine(directory.FullName, "oPenEfficiency.slnx"))) directory = directory.Parent;
            if (directory == null) throw new DirectoryNotFoundException("Repository root not found.");
            return directory.FullName;
        }

        private static XElement Load(string path) => XElement.Load(Path.Combine(Root(), path));
        private static XElement Tokens() => Load("oPenEfficiency/UI/Styles/Controls.xaml");
        private static double[] Thickness(string value)
        {
            var parts = value.Split(',').Select(number => double.Parse(number, CultureInfo.InvariantCulture)).ToArray();
            return parts.Length == 1 ? Enumerable.Repeat(parts[0], 4).ToArray() :
                parts.Length == 2 ? new[] { parts[0], parts[1], parts[0], parts[1] } : parts;
        }

        private static string Resource(XElement document, string key) => document.Descendants()
            .Single(element => (string)element.Attribute(Xaml + "Key") == key).Value.Trim();
        private static double[] Resolve(string expression)
        {
            if (expression.StartsWith("{StaticResource ", StringComparison.Ordinal))
                expression = Resource(Tokens(), expression.Substring(16).TrimEnd('}'));
            return Thickness(expression);
        }

        [TestMethod]
        public void TaskAndWorkspaceWindowEdgesMatchTheSettingsVerticalReference()
        {
            var settings = Load("oPenEfficiency/UI/Dialogs/SettingsWindow.xaml");
            var header = Thickness(Resource(settings, "SettingsHeaderMargin"));
            var footer = settings.Descendants().Single(element => (string)element.Attribute(Xaml + "Name") == "SettingsFooter");
            var footerPadding = Resolve((string)footer.Attribute("Padding"));
            foreach (var key in new[] { "WorkspaceOuterPadding", "CompactTaskOuterPadding" })
            {
                var edges = Thickness(Resource(Tokens(), key));
                Assert.AreEqual(header[1], edges[1], key + " top edge differs from Settings.");
                Assert.AreEqual(footerPadding[3], edges[3], key + " bottom edge differs from Settings.");
            }
            Assert.AreEqual(header[0], Thickness(Resource(Tokens(), "WorkspaceOuterPadding"))[0]);
        }

        [TestMethod]
        public void AgendaFormsHaveNoOutOfRangeRowsOrOverlappingCells()
        {
            var window = Load("oPenEfficiency/UI/Dialogs/SaveAgendaLayoutWindow.xaml");
            var expander = window.Descendants(Wpf + "Expander").Single(element => (string)element.Attribute("Header") == "LAYOUT SETTINGS");
            AssertGridCells(expander.Element(Wpf + "Grid"));
            var spacingGrid = expander.Descendants(Wpf + "TextBox").Single(element =>
                (string)element.Attribute(Xaml + "Name") == "TxtItemSpacing").Parent;
            AssertGridCells(spacingGrid);
        }

        private static void AssertGridCells(XElement grid)
        {
            var rowCount = grid.Element(Wpf + "Grid.RowDefinitions").Elements().Count();
            var occupied = new HashSet<string>();
            foreach (var child in grid.Elements().Where(element => !element.Name.LocalName.Contains(".")))
            {
                int row = (int?)child.Attribute("Grid.Row") ?? 0;
                int column = (int?)child.Attribute("Grid.Column") ?? 0;
                int span = (int?)child.Attribute("Grid.ColumnSpan") ?? 1;
                Assert.IsTrue(row >= 0 && row < rowCount, child.Name + " has an invalid row.");
                for (int c = column; c < column + span; c++)
                    Assert.IsTrue(occupied.Add(row + ":" + c), child.Name + " overlaps another control at " + row + ":" + c);
            }
        }

        [TestMethod]
        public void AgendaShellsUseSharedPaddingWithoutAddingAnOuterOverride()
        {
            foreach (var file in AgendaFiles)
            {
                var shell = Load(file).Elements(Wpf + "Border").Single();
                Assert.IsNull(shell.Attribute("Padding"), file + " duplicates the shared shell padding.");
                Assert.IsTrue(((string)shell.Attribute("Style")).Contains("Shell"), file);
            }
        }

        [TestMethod]
        public void AgendaUiVerticalGapsStayWithinTheSettingsGroupSpacingScale()
        {
            double maximum = Thickness(Resource(Tokens(), "Space12"))[0];
            foreach (var file in AgendaFiles)
            foreach (var attribute in Load(file).DescendantsAndSelf().Attributes().Where(attribute =>
                attribute.Name.LocalName == "Margin" || attribute.Name.LocalName == "Padding"))
            {
                if (attribute.Value.StartsWith("{Binding", StringComparison.Ordinal) || attribute.Value.StartsWith("{DynamicResource", StringComparison.Ordinal)) continue;
                var thickness = Resolve(attribute.Value);
                Assert.IsTrue(thickness[1] <= maximum && thickness[3] <= maximum,
                    file + " has excess vertical chrome spacing: " + attribute.Value);
            }
        }
    }
}
