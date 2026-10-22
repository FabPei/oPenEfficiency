using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI;

namespace oPenEfficiency.Command.Tests
{
    /// <summary>
    /// REG-MAGICRESIZER-GROUP-001: a grouped selection must keep Scale Selected Shapes clickable.
    /// The owner reported the command dead whenever the selection is a group; the production rule
    /// previously demanded two shapes, and a single group counts as one shape target.
    /// </summary>
    [TestClass]
    public sealed class MagicResizerSelectionRuleTests
    {
        [TestMethod]
        public void SingleShapeAndSingleGroupSelectionEnableScaleSelectedShapes()
        {
            FeatureDisplayInfo info = FeatureLibrary.GetFeatureInfo("BtnMagicResizer");
            Assert.IsTrue(info.MinSelection <= 1,
                "REG-MAGICRESIZER-GROUP-001: a single selected group must remain a clickable Scale Selected Shapes target, but MinSelection is " + info.MinSelection + ".");

            var rule = new SidebarFeatureStateRule
            {
                MinSelection = info.MinSelection,
                MaxSelection = info.MaxSelection,
                RequiredType = OfficeSelectionType.Shapes,
                RequiresTable = info.RequiresTable,
                RequiresHiddenShapes = info.RequiresHiddenShapes
            };

            // One selected group is one shape target and must satisfy the minimum.
            Assert.IsTrue(SidebarStateEvaluator.IsEnabled(rule,
                    new SidebarSelectionSnapshot { Type = OfficeSelectionType.Shapes, ShapeCount = 1 }),
                "One selected shape (or one group) must enable Scale Selected Shapes.");
            Assert.IsTrue(SidebarStateEvaluator.IsEnabled(rule,
                    new SidebarSelectionSnapshot { Type = OfficeSelectionType.Shapes, ShapeCount = 3 }),
                "Multiple selected shapes keep enabling Scale Selected Shapes.");
            Assert.IsFalse(SidebarStateEvaluator.IsEnabled(rule,
                    new SidebarSelectionSnapshot { Type = OfficeSelectionType.Shapes, ShapeCount = 0 }),
                "An empty shape selection stays disabled.");
        }
    }
}
