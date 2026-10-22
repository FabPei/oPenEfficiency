using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using oPenEfficiency.Models;
using oPenEfficiency.Utils;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Services
{
    internal sealed class CountryFlagInsertionService
    {
        private readonly PowerPointManager _manager;
        public CountryFlagInsertionService(PowerPointManager manager) { _manager = manager ?? throw new ArgumentNullException(nameof(manager)); }

        public bool Insert(CountryFlagEntry entry)
        {
            if (entry == null) return false;
            PowerPoint.Shape shape = null;
            try
            {
                string path = GetSvgPath(entry.AssetId);
                if (!File.Exists(path)) throw new FileNotFoundException("Bundled country flag is missing.", path);
                var app = _manager.GetApplication();
                var slide = app?.ActiveWindow?.View?.Slide as PowerPoint.Slide;
                if (app == null || slide == null) return false;
                app.StartNewUndoEntry();
                shape = slide.Shapes.AddPicture(path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 0, 0, -1, -1);
                shape.LockAspectRatio = Office.MsoTriState.msoTrue;
                const float targetWidth = 144f;
                if (shape.Width > targetWidth) shape.Width = targetWidth;
                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
                shape.Left = (slideWidth - shape.Width) / 2f;
                shape.Top = (slideHeight - shape.Height) / 2f;
                shape.Name = "oPE_Flag_" + entry.Alpha2;
                shape.AlternativeText = $"{entry.Name} flag ({entry.Alpha2})";
                shape.Select();
                Remember(entry.AssetId);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CountryFlagInsertionService.Insert");
                return false;
            }
            finally
            {
                if (shape != null && Marshal.IsComObject(shape)) Marshal.ReleaseComObject(shape);
            }
        }

        public static string GetSvgPath(string assetId) => Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "Resources", "Flags", assetId + ".svg");

        public static string GetThumbnailPath(string assetId) => Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "Resources", "Flags", "thumbs", assetId + ".png");

        private static void Remember(string assetId)
        {
            var config = JsonService.Load<SidebarConfig>("sidebar_config.json") ?? new SidebarConfig();
            var ids = config.RecentCountryFlagIds ?? new System.Collections.Generic.List<string>();
            ids.RemoveAll(id => string.Equals(id, assetId, StringComparison.OrdinalIgnoreCase));
            ids.Insert(0, assetId);
            config.RecentCountryFlagIds = ids.Take(8).ToList();
            JsonService.TrySave(config, "sidebar_config.json");
        }
    }
}
