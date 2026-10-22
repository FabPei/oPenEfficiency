using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Arrange selected shapes in a grid layout.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnArrangeGrid",
        Name = "Arrange Grid",
        Tooltip = "Arrange selected shapes in a grid",
        IconData = "M3,3H10V10H3V3M11,3H18V10H11V3M3,11H10V18H3V11M11,11H18V18H11V11Z",
        Color = "#F43F5E",
        Description = "Opens a grid layout dialog for arranging selected shapes with your chosen rows, columns, and spacing.",
        DetailedHelpText = "### Arrange Grid\n\nSelect two or more shapes, then choose positive row and column counts and a non-negative spacing. Press Enter to apply valid values or Escape to close without arranging. Right-click for direct two-, three-, four-, or five-column grids with automatic row calculation and 10-point spacing. Shapes follow selection order from the selection's top-left position; each column and row accommodates its largest shape. The operation can be undone as one step.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class ArrangeGridFeature
    {
        private static UI.ArrangeGridWindow _window;

        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var selected = manager.GetSelectedShapes();
            if (selected == null || selected.Count < 2) return false;

            if (_window != null && !_window.IsLoaded)
                _window = null;

            if (_window != null)
            {
                var existingApp = manager.GetApplication();
                if (existingApp != null) new System.Windows.Interop.WindowInteropHelper(_window).Owner = (IntPtr)existingApp.HWND;
                if (!_window.IsVisible) _window.Show();
                if (_window.WindowState == System.Windows.WindowState.Minimized)
                    _window.WindowState = System.Windows.WindowState.Normal;
                _window.Activate();
                _window.Focus();
                return true;
            }

            try
            {
                _window = new UI.ArrangeGridWindow(manager);
                var app = manager.GetApplication();
                if (app != null) new System.Windows.Interop.WindowInteropHelper(_window).Owner = (IntPtr)app.HWND;
                _window.Closed += (sender, args) => _window = null;
                _window.Show();
                if (!_window.IsLoaded || !_window.IsVisible)
                    throw new InvalidOperationException(
                        "Arrange Grid window closed before it became visible.");
                return true;
            }
            catch
            {
                try { _window?.Close(); }
                catch (Exception closeEx)
                {
                    ExceptionLogger.Log(closeEx, "ArrangeGridFeature.Execute.FailureCleanup");
                }
                _window = null;
                throw;
            }
        }

        public static bool Execute(PowerPointManager manager, int rows, int columns, float spacing = 10f)
        {
            if (manager == null || rows < 1 || columns < 1 || spacing < 0) return false;
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count < 2) return false;

            try
            {
                int count = shapeRange.Count;
                if ((long)rows * columns < count) return false;

                var originalPositions = new Tuple<Shape, float, float>[count];
                for (int i = 1; i <= count; i++) originalPositions[i - 1] = Tuple.Create(shapeRange[i], shapeRange[i].Left, shapeRange[i].Top);

                var geometry = Enumerable.Range(1, count)
                    .Select(index => new ShapeGeometry
                    {
                        Id = shapeRange[index].Id,
                        Left = shapeRange[index].Left,
                        Top = shapeRange[index].Top,
                        Width = shapeRange[index].Width,
                        Height = shapeRange[index].Height
                    })
                    .ToList();
                var plannedById = ShapeGeometryPlanner
                    .ArrangeGrid(geometry, rows, columns, spacing)
                    .ToDictionary(item => item.Id);

                try { manager.GetApplication().StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ArrangeGridFeature.StartNewUndoEntry"); }

                try
                {
                    for (int index = 0; index < count; index++)
                    {
                        var shape = shapeRange[index + 1];
                        shape.Left = plannedById[shape.Id].Left;
                        shape.Top = plannedById[shape.Id].Top;
                    }
                }
                catch
                {
                    foreach (var original in originalPositions)
                    {
                        try { original.Item1.Left = original.Item2; original.Item1.Top = original.Item3; } catch { /* Preserve the primary layout error if one stale shape cannot be restored. */ }
                    }
                    throw;
                }

                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "ArrangeGridFeature.Execute");
                return false;
            }
        }

        public static bool ExecuteColumns(PowerPointManager manager, int columns, float spacing = 10f)
        {
            var shapes = manager?.GetSelectedShapes();
            if (shapes == null || shapes.Count < 2 || columns < 1) return false;
            int rows = (int)Math.Ceiling(shapes.Count / (double)columns);
            return Execute(manager, rows, columns, spacing);
        }
    }
}
