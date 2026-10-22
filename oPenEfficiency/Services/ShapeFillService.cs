using System;
using System.Collections.Generic;
using System.Linq;
using oPenEfficiency.Models;
using oPenEfficiency.Utils;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Services
{
    internal sealed class ShapeFillService
    {
        private readonly PowerPointManager _manager;
        public ShapeFillService(PowerPointManager manager) { _manager = manager ?? throw new ArgumentNullException(nameof(manager)); }

        public bool ApplyRgb(int bgr)
        {
            bool success = Apply(fill =>
            {
                fill.Visible = Office.MsoTriState.msoTrue;
                fill.Solid();
                fill.ForeColor.ObjectThemeColor = Office.MsoThemeColorIndex.msoNotThemeColor;
                fill.ForeColor.RGB = bgr;
            });
            if (success) RememberRecent(bgr);
            return success;
        }

        public bool ApplyTheme(Office.MsoThemeColorIndex color, float tint)
        {
            return Apply(fill =>
            {
                fill.Visible = Office.MsoTriState.msoTrue;
                fill.Solid();
                fill.ForeColor.ObjectThemeColor = color;
                fill.ForeColor.TintAndShade = tint;
            });
        }

        public bool ApplyNoFill() => Apply(fill => fill.Visible = Office.MsoTriState.msoFalse);

        private bool Apply(Action<PowerPoint.FillFormat> mutation)
        {
            try
            {
                var shapes = _manager.GetSelectedShapes();
                if (shapes == null || shapes.Count == 0) return false;
                _manager.GetApplication()?.StartNewUndoEntry();
                foreach (PowerPoint.Shape shape in shapes) mutation(shape.Fill);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ShapeFillService.Apply");
                return false;
            }
        }

        private static void RememberRecent(int bgr)
        {
            var color = System.Drawing.ColorTranslator.FromOle(bgr);
            string hex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
            var config = JsonService.Load<SidebarConfig>("sidebar_config.json") ?? new SidebarConfig();
            var colors = config.RecentShapeFillColors ?? new List<string>();
            colors.RemoveAll(value => string.Equals(value, hex, StringComparison.OrdinalIgnoreCase));
            colors.Insert(0, hex);
            config.RecentShapeFillColors = colors.Take(10).ToList();
            JsonService.TrySave(config, "sidebar_config.json");
        }
    }
}
