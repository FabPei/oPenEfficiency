using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using oPenEfficiency.Features;
using oPenEfficiency.Utils;
using oPenEfficiency.UI;
using oPenEfficiency.Services;

namespace oPenEfficiency
{
    public partial class ThisAddIn
    {
        private oPenEfficiency.UI.RatingToolbar _ratingToolbar;

        private void ShowStarRatingToolbar(string altText)
        {
            try
            {
                var manager = new PowerPointManager(this.Application);
                _ratingToolbar = _modelessUi.ReplaceWindow(
                    "StarRating",
                    () => new oPenEfficiency.UI.RatingToolbar(manager, altText),
                    window => oPenEfficiency.UI.Lifecycle.ModelessUiCoordinator.PositionAtCursor(
                        window, 10, 10, _dpiX, _dpiY));
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ShowStarRatingToolbar");
            }
        }

        private oPenEfficiency.UI.ProgressSeriesToolbar _progressSeriesToolbar;

        private void ShowProgressSeriesToolbar(PowerPoint.Shape shape)
        {
            try
            {
                var manager = new PowerPointManager(this.Application);
                _progressSeriesToolbar = _modelessUi.GetOrCreateWindow(
                    "ProgressSeries",
                    () => new oPenEfficiency.UI.ProgressSeriesToolbar(manager, shape),
                    window => window.UpdateTarget(shape),
                    window => oPenEfficiency.UI.Lifecycle.ModelessUiCoordinator.PositionAtShape(
                        window, this.Application, shape, 10, 0, _dpiX, _dpiY));
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ShowProgressSeriesToolbar");
                System.Windows.MessageBox.Show($"Error showing Progress Series toolbar: {ex.Message}");
            }
        }

        private oPenEfficiency.UI.NumerationToolbar _numerationToolbar;

        private void ShowNumerationToolbar(PowerPoint.Shape shape)
        {
            try
            {
                var manager = new PowerPointManager(this.Application);
                _numerationToolbar = _modelessUi.ReplaceWindow(
                    "Numeration",
                    () => new oPenEfficiency.UI.NumerationToolbar(manager, shape),
                    window => oPenEfficiency.UI.Lifecycle.ModelessUiCoordinator.PositionAtCursor(
                        window, -210, 20, _dpiX, _dpiY));
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ShowNumerationToolbar");
            }
        }

        private void ShowExcelValueToolbar(PowerPoint.Shape shape)
        {
            try
            {
                var manager = new PowerPointManager(this.Application);
                var toolbar = _modelessUi.TrackForm("ExcelValue", new oPenEfficiency.UI.FloatingToolbar());

                toolbar.AddLabel("🔗");

                toolbar.AddButton("Refresh", null, (s, e) =>
                {
                    oPenEfficiency.Features.Utilities.InsertExcelValueFeature.RefreshShape(manager, shape);
                    toolbar.Close();
                });

                toolbar.AddSeparator();

                toolbar.AddDropDownButton("Format", null, (dropDown) =>
                {
                    oPenEfficiency.Features.Utilities.ExcelValueLink.TryParse(shape.AlternativeText ?? "", out var link);
                    int curDec = link != null ? link.Decimals : 0;
                    bool curThou = link != null && link.Thousands;
                    bool curPct = link != null && link.Percent;

                    for (int dCount = 0; dCount <= 3; dCount++)
                    {
                        int dc = dCount;
                        var item = dropDown.DropDownItems.Add(dc + " decimal places");
                        item.Click += (s, e) =>
                        {
                            oPenEfficiency.Features.Utilities.InsertExcelValueFeature.SetFormat(shape, dc, curThou, curPct);
                            toolbar.Close();
                        };
                    }

                    dropDown.DropDownItems.Add(new System.Windows.Forms.ToolStripSeparator());

                    var thouItem = new System.Windows.Forms.ToolStripMenuItem("Thousands Separator") { Checked = curThou };
                    thouItem.Click += (s, e) =>
                    {
                        oPenEfficiency.Features.Utilities.InsertExcelValueFeature.SetFormat(shape, Math.Max(curDec, 0), !curThou, curPct);
                        toolbar.Close();
                    };
                    dropDown.DropDownItems.Add(thouItem);

                    var pctItem = new System.Windows.Forms.ToolStripMenuItem("Percent") { Checked = curPct };
                    pctItem.Click += (s, e) =>
                    {
                        oPenEfficiency.Features.Utilities.InsertExcelValueFeature.SetFormat(shape, Math.Max(curDec, 0), curThou, !curPct);
                        toolbar.Close();
                    };
                    dropDown.DropDownItems.Add(pctItem);
                });

                var mousePos = System.Windows.Forms.Cursor.Position;
                toolbar.Location = new System.Drawing.Point(mousePos.X, mousePos.Y + 20);
                toolbar.Show();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ShowExcelValueToolbar");
            }
        }

        private void ShowTableLinkToolbar(PowerPoint.Shape shape)
        {
            try
            {
                var manager = new PowerPointManager(this.Application);
                var toolbar = _modelessUi.TrackForm("TableLink", new oPenEfficiency.UI.FloatingToolbar());

                oPenEfficiency.Features.Utilities.TableLink.TryParse(shape.AlternativeText ?? "", out var link);
                if (link != null)
                {
                    string fileName = string.IsNullOrEmpty(link.Path)
                        ? "?" : System.IO.Path.GetFileName(link.Path);
                    toolbar.AddLabel($"🔗 {fileName}  [{link.Sheet}!{link.Range}]");
                    toolbar.AddSeparator();
                }
                else
                {
                    toolbar.AddLabel("🔗");
                }

                toolbar.AddButton("Refresh", null, (s, e) =>
                {
                    oPenEfficiency.Features.Utilities.TableLinkFeature.RefreshTable(manager, shape);
                    toolbar.Close();
                });

                toolbar.AddButton("Change Range", null, (s, e) =>
                {
                    oPenEfficiency.Features.Utilities.TableLinkFeature.ChangeRange(manager, shape);
                    toolbar.Close();
                });

                toolbar.AddButton("Relink Source", null, (s, e) =>
                {
                    oPenEfficiency.Features.Utilities.TableLinkFeature.RelinkSource(manager, shape);
                    toolbar.Close();
                });

                toolbar.AddButton("Auto-Resize", null, (s, e) =>
                {
                    oPenEfficiency.Features.Utilities.TableLinkFeature.ToggleAutoResize(shape);
                    toolbar.Close();
                }, isChecked: link != null && link.AutoResize);

                toolbar.AddButton("Open Source", null, (s, e) =>
                {
                    oPenEfficiency.Features.Utilities.TableLinkFeature.OpenSource(shape);
                    toolbar.Close();
                });

                toolbar.AddSeparator();

                toolbar.AddButton("Unlink", null, (s, e) =>
                {
                    var result = System.Windows.MessageBox.Show(
                        "Remove the Excel source connection from this table?\n\nThe current PowerPoint table content will be kept.",
                        "Unlink Excel Table",
                        System.Windows.MessageBoxButton.OKCancel,
                        System.Windows.MessageBoxImage.Warning);
                    if (result == System.Windows.MessageBoxResult.OK)
                        oPenEfficiency.Features.Utilities.TableLinkFeature.Unlink(shape);
                    toolbar.Close();
                });

                var mousePos = System.Windows.Forms.Cursor.Position;
                toolbar.Location = new System.Drawing.Point(mousePos.X, mousePos.Y + 20);
                toolbar.Show();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ShowTableLinkToolbar");
            }
        }

        private void ShowInlineExcelValueToolbar(PowerPoint.Shape shape)
        {
            try
            {
                var manager = new PowerPointManager(this.Application);
                var toolbar = _modelessUi.TrackForm("InlineExcelValue", new oPenEfficiency.UI.FloatingToolbar());

                toolbar.AddLabel("🔗");

                toolbar.AddButton("Refresh", null, (s, e) =>
                {
                    var selection = this.Application?.ActiveWindow?.Selection;
                    if (!oPenEfficiency.Features.Utilities.InsertExcelValueFeature.RefreshInlineAtSelection(manager, selection))
                    {
                        oPenEfficiency.Features.Utilities.InsertExcelValueFeature.RefreshInline(manager, shape);
                    }
                    toolbar.Close();
                });

                toolbar.AddSeparator();

                toolbar.AddButton("Delete Value", null, (s, e) =>
                {
                    var result = System.Windows.MessageBox.Show(
                        "Delete this linked value from the text?",
                        "Delete Linked Excel Value",
                        System.Windows.MessageBoxButton.OKCancel,
                        System.Windows.MessageBoxImage.Warning);
                    if (result == System.Windows.MessageBoxResult.OK)
                    {
                        var selection = this.Application?.ActiveWindow?.Selection;
                        oPenEfficiency.Features.Utilities.InsertExcelValueFeature.RemoveInlineAtSelection(selection);
                    }
                    toolbar.Close();
                });

                toolbar.AddButton("Unlink", null, (s, e) =>
                {
                    var result = System.Windows.MessageBox.Show(
                        "Remove the Excel source connection from this value?\n\nThe displayed text will be kept.",
                        "Unlink Excel Value",
                        System.Windows.MessageBoxButton.OKCancel,
                        System.Windows.MessageBoxImage.Warning);
                    if (result == System.Windows.MessageBoxResult.OK)
                    {
                        var selection = this.Application?.ActiveWindow?.Selection;
                        oPenEfficiency.Features.Utilities.InsertExcelValueFeature.UnlinkInlineAtSelection(selection);
                    }
                    toolbar.Close();
                });

                toolbar.AddSeparator();

                toolbar.AddDropDownButton("Format", null, (dropDown) =>
                {
                    var selection = this.Application?.ActiveWindow?.Selection;
                    var link = oPenEfficiency.Features.Utilities.InsertExcelValueFeature.GetInlineAtSelection(selection)
                        ?? oPenEfficiency.Features.Utilities.InsertExcelValueFeature.GetFirstInline(shape);
                    int curDec = link != null ? link.Decimals : 0;
                    bool curThou = link != null && link.Thousands;
                    bool curPct = link != null && link.Percent;

                    for (int dCount = 0; dCount <= 3; dCount++)
                    {
                        int dc = dCount;
                        var item = dropDown.DropDownItems.Add(dc + " decimal places");
                        item.Click += (s, e) =>
                        {
                            var currentSelection = this.Application?.ActiveWindow?.Selection;
                            if (!oPenEfficiency.Features.Utilities.InsertExcelValueFeature.SetFormatInlineAtSelection(currentSelection, dc, curThou, curPct))
                            {
                                oPenEfficiency.Features.Utilities.InsertExcelValueFeature.SetFormatInline(shape, dc, curThou, curPct);
                            }
                            toolbar.Close();
                        };
                    }

                    dropDown.DropDownItems.Add(new System.Windows.Forms.ToolStripSeparator());

                    var thouItem = new System.Windows.Forms.ToolStripMenuItem("Thousands Separator") { Checked = curThou };
                    thouItem.Click += (s, e) =>
                    {
                        var currentSelection = this.Application?.ActiveWindow?.Selection;
                        if (!oPenEfficiency.Features.Utilities.InsertExcelValueFeature.SetFormatInlineAtSelection(currentSelection, Math.Max(curDec, 0), !curThou, curPct))
                        {
                            oPenEfficiency.Features.Utilities.InsertExcelValueFeature.SetFormatInline(shape, Math.Max(curDec, 0), !curThou, curPct);
                        }
                        toolbar.Close();
                    };
                    dropDown.DropDownItems.Add(thouItem);

                    var pctItem = new System.Windows.Forms.ToolStripMenuItem("Percent") { Checked = curPct };
                    pctItem.Click += (s, e) =>
                    {
                        var currentSelection = this.Application?.ActiveWindow?.Selection;
                        if (!oPenEfficiency.Features.Utilities.InsertExcelValueFeature.SetFormatInlineAtSelection(currentSelection, Math.Max(curDec, 0), curThou, !curPct))
                        {
                            oPenEfficiency.Features.Utilities.InsertExcelValueFeature.SetFormatInline(shape, Math.Max(curDec, 0), curThou, !curPct);
                        }
                        toolbar.Close();
                    };
                    dropDown.DropDownItems.Add(pctItem);
                });

                var mousePos = System.Windows.Forms.Cursor.Position;
                toolbar.Location = new System.Drawing.Point(mousePos.X, mousePos.Y + 20);
                toolbar.Show();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ShowInlineExcelValueToolbar");
            }
        }

        private void ShowThermometerToolbar(string altText)
        {
            try
            {
                var manager = new PowerPointManager(this.Application);
                var toolbar = _modelessUi.TrackForm("Thermometer", new oPenEfficiency.UI.FloatingToolbar());

                // Parse state
                float percentage = 0.5f;
                int colorRgb = 0x0000FF; // Default

                var parts = altText.Split('|');
                if (parts.Length >= 2) float.TryParse(parts[1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out percentage);
                if (parts.Length >= 3) int.TryParse(parts[2], out colorRgb);

                toolbar.AddLabel("🌡️");

                // 1. Percentage Slider
                toolbar.AddSlider(0, 100, (int)(percentage * 100), (s, e) =>
                {
                    var track = s as System.Windows.Forms.TrackBar;
                    if (track != null)
                    {
                        float p = track.Value / 100f;
                        ThermometerFeature.Execute(manager, p, colorRgb);
                    }
                });

                toolbar.AddSeparator();

                // 2. Color Dropdown
                toolbar.AddDropDownButton("Color", null, (dropDown) =>
                {
                    int[] colors = new[] { 0x0000FF, 0x00FF00, 0xFF0000, 0x008CFF, 0x00FFFF };
                    string[] names = new[] { "Red", "Green", "Blue", "Orange", "Cyan" };

                    for (int i = 0; i < colors.Length; i++)
                    {
                        int c = colors[i];
                        var item = dropDown.DropDownItems.Add(names[i]);
                        item.Click += (s, e) =>
                        {
                            ThermometerFeature.Execute(manager, percentage, c);
                            toolbar.Close();
                        };
                    }

                    dropDown.DropDownItems.Add(new System.Windows.Forms.ToolStripSeparator());
                    var custom = dropDown.DropDownItems.Add("Custom...");
                    custom.Click += (s, e) =>
                    {
                        var dialog = new System.Windows.Forms.ColorDialog();
                        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            var color = dialog.Color;
                            int rgb = color.R | (color.G << 8) | (color.B << 16);
                            ThermometerFeature.Execute(manager, percentage, rgb);
                            toolbar.Close();
                        }
                    };
                });

                var mousePos = System.Windows.Forms.Cursor.Position;
                toolbar.Location = new System.Drawing.Point(mousePos.X + 10, mousePos.Y + 10);
                toolbar.Show();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Toolbar Error: " + ex.Message);
            }
        }

        private void UpdateAdornerPosition()
        {
            if (_thermometerAdorner == null || _trackedShape == null) return;
            try
            {
                var activeWin = this.Application.ActiveWindow;
                // Get screen coordinates of the shape
                float shapeLeft = _trackedShape.Left;
                float shapeTop = _trackedShape.Top;
                float shapeWidth = _trackedShape.Width;

                int screenX = activeWin.PointsToScreenPixelsX(shapeLeft + (shapeWidth / 2));
                int screenY = activeWin.PointsToScreenPixelsY(shapeTop);

                double adornerWidth = _thermometerAdorner.Width > 0 ? _thermometerAdorner.Width : 300;
                double adornerHeight = _thermometerAdorner.Height > 0 ? _thermometerAdorner.Height : 48;

                _thermometerAdorner.Left = screenX - (adornerWidth / 2);
                _thermometerAdorner.Top = screenY - adornerHeight - 20;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.UpdateAdornerPosition");
                CloseThermometerAdorner();
            }
        }

        private void CloseThermometerAdorner()
        {
            if (_thermometerAdorner != null)
            {
                try { _thermometerAdorner.Close(); } catch (Exception ex) { ExceptionLogger.Log(ex, "ThisAddIn.CloseThermometerAdorner"); }
                _thermometerAdorner = null;
            }
            _trackedShape = null;

        }

        private void ShowTrafficLightMenu(string altText)
        {
            try
            {
                var manager = new PowerPointManager(this.Application);
                _modelessUi.ReplaceWindow(
                    "TrafficLight",
                    () =>
                    {
                        var window = new oPenEfficiency.UI.TrafficLightMenu(manager, altText);
                        new System.Windows.Interop.WindowInteropHelper(window).Owner = (IntPtr)Application.HWND;
                        return window;
                    },
                    window => oPenEfficiency.UI.Lifecycle.ModelessUiCoordinator.PositionAtCursor(
                        window, 10, 10, _dpiX, _dpiY));
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ShowTrafficLightMenu");
            }
        }

        private oPenEfficiency.UI.CheckboxToolbar _checkboxToolbar;

        private void ShowCheckboxMenu(string altText)
        {
            try
            {
                var manager = new PowerPointManager(this.Application);
                _checkboxToolbar = _modelessUi.ReplaceWindow(
                    "Checkbox",
                    () => new oPenEfficiency.UI.CheckboxToolbar(manager, altText),
                    null);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ThisAddIn.ShowCheckboxMenu");
            }
        }

        private void ShowHarveyBallMenu()
        {
            var menu = new System.Windows.Forms.ContextMenuStrip();
            var levels = new[]
            {
                new { Name = "0%", Value = 0f },
                new { Name = "10%", Value = 0.1f },
                new { Name = "25%", Value = 0.25f },
                new { Name = "50%", Value = 0.5f },
                new { Name = "75%", Value = 0.75f },
                new { Name = "100%", Value = 1f }
            };

            foreach (var level in levels)
            {
                var item = menu.Items.Add(level.Name);
                float val = level.Value;
                item.Click += (s, args) =>
                {
                    var manager = new PowerPointManager(this.Application);
                    oPenEfficiency.Features.HarveyBallFeature.Execute(manager, val);
                };
            }

            menu.Items.Add(new System.Windows.Forms.ToolStripSeparator());
            var cycleItem = menu.Items.Add("Cycle (+25%)");
            cycleItem.Click += (s, args) =>
            {
                // Re-use logic from MainSidebar if possible, or duplicate for now
                var manager = new PowerPointManager(this.Application);
                CycleHarveyBall(manager);
            };

            var mousePos = System.Windows.Forms.Cursor.Position;
            menu.Show(mousePos);
        }

        private void CycleHarveyBall(PowerPointManager manager)
        {
             var selection = manager.GetSelectedShapes();
             float nextPercentage = 0.25f;

             if (selection != null && selection.Count == 1)
             {
                 var selected = selection[1];
                 string alt = selected.AlternativeText;
                 if (alt.StartsWith("oPE_HarveyBall"))
                 {
                     var parts = alt.Split('|');
                     if (parts.Length == 2 && float.TryParse(parts[1], out float current))
                     {
                         nextPercentage = current + 0.25f;
                         if (nextPercentage > 1.01f) nextPercentage = 0f;
                     }
                 }
             }
             oPenEfficiency.Features.HarveyBallFeature.Execute(manager, nextPercentage);
        }

    }
}
