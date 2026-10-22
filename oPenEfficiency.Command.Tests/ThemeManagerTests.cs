using System.Windows;
using System.Windows.Media;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Command.Tests
{
    [TestClass]
    public sealed class ThemeManagerTests
    {
        [STATestMethod]
        public void SidebarBackgroundResourceTracksLightDarkLightSwitch()
        {
            var app = Application.Current ?? new Application
            {
                ShutdownMode = ShutdownMode.OnExplicitShutdown
            };
            string originalTheme = ThemeManager.GetCurrentTheme();

            try
            {
                AssertSidebarColor(app, "Light", Color.FromRgb(0xF3, 0xF4, 0xF6));
                AssertSidebarColor(app, "Dark", Color.FromRgb(0x29, 0x29, 0x29));
                AssertSidebarColor(app, "Light", Color.FromRgb(0xF3, 0xF4, 0xF6));
            }
            finally
            {
                ThemeManager.ApplyTheme(originalTheme);
            }
        }

        private static void AssertSidebarColor(Application app, string theme, Color expected)
        {
            ThemeManager.ApplyTheme(theme);

            var brush = app.TryFindResource("SidebarBackgroundBrush") as SolidColorBrush;
            Assert.IsNotNull(brush, theme + " theme has no SidebarBackgroundBrush.");
            Assert.AreEqual(expected, brush.Color, theme + " Sidebar background color is incorrect.");
        }
    }
}
