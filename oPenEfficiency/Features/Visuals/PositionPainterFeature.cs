using System;
using oPenEfficiency.Services.Attributes;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Visuals
{
    [FeatureMetadata(
        Id = "BtnPositionPainter",
        Name = "Copy/Paste Position & Size",
        SearchAliases = new[] { "Position Painter", "Paint Position (Left & Top)", "Paint Size (Width & Height)", "Paint Position and Size" },
        Tooltip = "Copy selected position and size properties to other shapes",
        IconData = "M17,12V3H7V12H17M14,10H10V5H14V10M20,7V14A2,2 0 0,1 18,16H7V14H18V7H20M5,12V21H7V12H5Z",
        Color = "#10B981",
        Description = "Copies the exact position and dimensions of one shape and pastes chosen properties onto other shapes in the same presentation.",
        DetailedHelpText = "### Copy/Paste Position & Size\n\nSelect exactly one source shape to copy its position and size and open the tool. Choose Left, Top, Width, and/or Height, select one or more target shapes in the same presentation, and click Paste. Right-click presets the tool for position only, size only, or both. Copy Again replaces the source values with the current single-shape selection. Target aspect-ratio locks are preserved, partial failures are rolled back, and each successful Paste is one PowerPoint Undo step.",
        MinSelection = 1,
        MaxSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class PositionPainterFeature
    {
        private static oPenEfficiency.UI.PositionPainterWindow _window;

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, true, true, true, true);
        }

        public static bool Execute(
            PowerPointManager manager,
            bool applyLeft,
            bool applyTop,
            bool applyWidth,
            bool applyHeight)
        {
            if (manager == null || manager.GetSelectedShapes()?.Count != 1) return false;
            try
            {
                if (_window != null && _window.IsVisible)
                {
                    _window.ConfigureProperties(applyLeft, applyTop, applyWidth, applyHeight);
                    _window.Activate();
                    return true;
                }

                _window = new oPenEfficiency.UI.PositionPainterWindow(
                    manager,
                    applyLeft,
                    applyTop,
                    applyWidth,
                    applyHeight);
                var app = manager.GetApplication();
                if (app != null)
                    new System.Windows.Interop.WindowInteropHelper(_window).Owner = new IntPtr(app.HWND);
                _window.Closed += (sender, args) => _window = null;
                _window.Show();
                return true;
            }
            catch (Exception ex)
            {
                var failedWindow = _window;
                _window = null;
                ExceptionLogger.Log(ex, "PositionPainterFeature.Execute");
                try { failedWindow?.Close(); }
                catch (Exception closeEx)
                {
                    ExceptionLogger.Log(closeEx, "PositionPainterFeature.Execute.FailureCleanup");
                }
                throw;
            }
        }
    }
}
