using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;
using oPenEfficiency.UI;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class FeatureIconColorPolicyTests
    {
        [TestMethod]
        public void SidebarGlyphRestoresTheFourDipOpticalExpansion()
        {
            Assert.AreEqual(20d, FeatureIconSizing.ResolveGlyphSize(16d), 0.001d);
        }

        [DataTestMethod]
        [DataRow(FeatureIconAccentRole.ActionBlue, FeatureIconColorPolicy.AccentBlueResourceKey)]
        [DataRow(FeatureIconAccentRole.CreateGreen, FeatureIconColorPolicy.AccentGreenResourceKey)]
        [DataRow(FeatureIconAccentRole.SuccessGreen, FeatureIconColorPolicy.SuccessGreenResourceKey)]
        [DataRow(FeatureIconAccentRole.AttentionYellow, FeatureIconColorPolicy.AccentYellowResourceKey)]
        [DataRow(FeatureIconAccentRole.DestructiveRed, FeatureIconColorPolicy.AccentRedResourceKey)]
        [DataRow(FeatureIconAccentRole.ContentColor, FeatureIconColorPolicy.AccentBlueResourceKey)]
        public void AccentLayerUsesAuditedSemanticResource(
            FeatureIconAccentRole accentRole,
            string expectedResourceKey)
        {
            FeatureIconColorDecision decision = FeatureIconColorPolicy.ResolveForLayer(
                IconColorModes.PowerPoint,
                FeatureIconVisualState.Normal,
                "#F43F5E",
                isAccent: true,
                accentRole: accentRole);

            Assert.IsFalse(decision.UsesStoredFeatureColor);
            Assert.AreEqual(expectedResourceKey, decision.ResourceKey);
        }

        [TestMethod]
        public void SuccessAndCreateGreenUseIndependentSemanticResources()
        {
            FeatureIconColorDecision create = FeatureIconColorPolicy.ResolveForLayer(
                IconColorModes.PowerPoint,
                FeatureIconVisualState.Normal,
                null,
                isAccent: true,
                accentRole: FeatureIconAccentRole.CreateGreen);
            FeatureIconColorDecision success = FeatureIconColorPolicy.ResolveForLayer(
                IconColorModes.PowerPoint,
                FeatureIconVisualState.Normal,
                null,
                isAccent: true,
                accentRole: FeatureIconAccentRole.SuccessGreen);

            Assert.AreEqual(FeatureIconColorPolicy.AccentGreenResourceKey, create.ResourceKey);
            Assert.AreEqual(FeatureIconColorPolicy.SuccessGreenResourceKey, success.ResourceKey);
            Assert.AreNotEqual(create.ResourceKey, success.ResourceKey);
        }

        [TestMethod]
        public void PowerPointBaseLayerCannotUseStoredFeatureColor()
        {
            FeatureIconColorDecision decision = FeatureIconColorPolicy.ResolveForLayer(
                IconColorModes.PowerPoint,
                FeatureIconVisualState.Normal,
                "#F43F5E",
                isAccent: false,
                accentRole: FeatureIconAccentRole.DestructiveRed);

            Assert.IsFalse(decision.UsesStoredFeatureColor);
            Assert.AreEqual(FeatureIconColorPolicy.NormalResourceKey, decision.ResourceKey);
        }

        [TestMethod]
        public void MissingAndInvalidModesResolveToPowerPoint()
        {
            Assert.AreEqual(IconColorModes.PowerPoint, IconColorModes.Normalize(null));
            Assert.AreEqual(IconColorModes.PowerPoint, IconColorModes.Normalize("unknown"));

            FeatureIconColorDecision decision = FeatureIconColorPolicy.ResolveForLayer(
                "unknown",
                FeatureIconVisualState.Normal,
                "#F43F5E",
                isAccent: false,
                accentRole: FeatureIconAccentRole.None);

            Assert.IsFalse(decision.UsesStoredFeatureColor);
            Assert.AreEqual(FeatureIconColorPolicy.NormalResourceKey, decision.ResourceKey);
        }

        [TestMethod]
        public void LegacyConfigWithoutModeUsesPowerPointFallback()
        {
            var serializer = new DataContractJsonSerializer(typeof(SidebarConfig));
            SidebarConfig legacy;
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("{}")))
            {
                legacy = (SidebarConfig)serializer.ReadObject(stream);
            }

            Assert.AreEqual(IconColorModes.PowerPoint, IconColorModes.Normalize(legacy.IconColorMode));
        }

        [TestMethod]
        public void PersistedFeatureColorsModeRoundTrips()
        {
            var config = new SidebarConfig { IconColorMode = IconColorModes.FeatureColors };
            var serializer = new DataContractJsonSerializer(typeof(SidebarConfig));

            SidebarConfig restored;
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, config);
                string json = Encoding.UTF8.GetString(stream.ToArray());
                StringAssert.Contains(json, "IconColorMode");
                stream.Position = 0;
                restored = (SidebarConfig)serializer.ReadObject(stream);
            }

            Assert.AreEqual(IconColorModes.FeatureColors, restored.IconColorMode);
        }

        [TestMethod]
        public void FeatureColorsApplyStoredColorToBothLayers()
        {
            const string storedColor = "#F43F5E";

            FeatureIconColorDecision baseLayer = FeatureIconColorPolicy.ResolveForLayer(
                IconColorModes.FeatureColors,
                FeatureIconVisualState.Normal,
                storedColor,
                isAccent: false,
                accentRole: FeatureIconAccentRole.None);
            FeatureIconColorDecision accentLayer = FeatureIconColorPolicy.ResolveForLayer(
                IconColorModes.FeatureColors,
                FeatureIconVisualState.Normal,
                storedColor,
                isAccent: true,
                accentRole: FeatureIconAccentRole.DestructiveRed);

            Assert.IsTrue(baseLayer.UsesStoredFeatureColor);
            Assert.IsTrue(accentLayer.UsesStoredFeatureColor);
            Assert.AreEqual(storedColor, baseLayer.StoredFeatureColor);
            Assert.AreEqual(storedColor, accentLayer.StoredFeatureColor);
        }

        [TestMethod]
        public void DisabledIconsUseMutedThemeResourceInBothModes()
        {
            foreach (string mode in new[] { IconColorModes.PowerPoint, IconColorModes.FeatureColors })
            {
                foreach (bool isAccent in new[] { false, true })
                {
                    FeatureIconColorDecision decision = FeatureIconColorPolicy.ResolveForLayer(
                        mode,
                        FeatureIconVisualState.Disabled,
                        "#3B82F6",
                        isAccent,
                        FeatureIconAccentRole.CreateGreen);

                    Assert.IsFalse(decision.UsesStoredFeatureColor);
                    Assert.AreEqual(FeatureIconColorPolicy.DisabledResourceKey, decision.ResourceKey);
                }
            }
        }
    }
}
