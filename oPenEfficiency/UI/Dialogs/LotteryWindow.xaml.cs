using System;
using System.Collections.Generic;
using System.Windows;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI
{
    public partial class LotteryWindow : Window
    {
        private PowerPointManager _pptManager;
        private List<string> _previousWinnersIds = new List<string>();
        private List<string> _currentSelectionIds = new List<string>();
        private List<string> _drawnArrowIds = new List<string>();

        public LotteryWindow(PowerPointManager pptManager)
        {
            InitializeComponent();
            _pptManager = pptManager;
            UpdateSelectionState();
        }

        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
                this.DragMove();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            RemoveArrows();
            Close();
        }

        private void BtnPickWinners_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selection = _pptManager.GetSelectedShapes();
                if (selection == null || selection.Count == 0)
                {
                    MessageBox.Show("Please select objects on the slide to pick winners from.", "No Selection", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                int count = 1;
                if (!int.TryParse(TxtWinnerCount.Text, out count) || count <= 0)
                {
                    count = 1;
                    TxtWinnerCount.Text = "1";
                }

                // Check if selection changed
                List<string> newSelectionIds = new List<string>();
                for (int i = 1; i <= selection.Count; i++)
                {
                    newSelectionIds.Add(selection[i].Id.ToString());
                }

                bool selectionChanged = false;
                if (newSelectionIds.Count != _currentSelectionIds.Count)
                {
                    selectionChanged = true;
                }
                else
                {
                    foreach (var id in newSelectionIds)
                    {
                        if (!_currentSelectionIds.Contains(id))
                        {
                            selectionChanged = true;
                            break;
                        }
                    }
                }

                if (selectionChanged)
                {
                    _previousWinnersIds.Clear();
                    _currentSelectionIds = newSelectionIds;
                }

                List<Shape> eligibleCandidates = new List<Shape>();
                for (int i = 1; i <= selection.Count; i++)
                {
                    var shp = selection[i];
                    bool isArrow = _drawnArrowIds.Contains(shp.Id.ToString());
                    if (isArrow) continue;

                    if (ChkExcludePrevious.IsChecked == true)
                    {
                        if (_previousWinnersIds.Contains(shp.Id.ToString())) continue;
                    }
                    eligibleCandidates.Add(shp);
                }

                if (eligibleCandidates.Count == 0)
                {
                    MessageBox.Show("No eligible candidates remaining! Please expand your selection or uncheck 'Exclude previous winners'.", "Lottery", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int pickCount = Math.Min(count, eligibleCandidates.Count);
                List<Shape> winners = new List<Shape>();
                Random rnd = new Random();

                for (int i = 0; i < pickCount; i++)
                {
                    int index = rnd.Next(eligibleCandidates.Count);
                    var winner = eligibleCandidates[index];
                    winners.Add(winner);
                    eligibleCandidates.RemoveAt(index);
                    if (!_previousWinnersIds.Contains(winner.Id.ToString()))
                    {
                        _previousWinnersIds.Add(winner.Id.ToString());
                    }
                }

                DrawArrows(winners);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while picking winners: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateSelectionState()
        {
            try
            {
                var selection = _pptManager.GetSelectedShapes();
                if (selection != null)
                {
                    _currentSelectionIds.Clear();
                    for (int i = 1; i <= selection.Count; i++)
                    {
                        _currentSelectionIds.Add(selection[i].Id.ToString());
                    }
                }
            }
            catch { }
        }

        private void DrawArrows(List<Shape> winners)
        {
            try
            {
                var app = _pptManager.GetApplication();
                if (app.ActiveWindow == null || app.ActiveWindow.View.Slide == null) return;
                var slide = (Slide)app.ActiveWindow.View.Slide;

                foreach (var winner in winners)
                {
                    // Draw a red down-arrow pointing to the top center of the shape
                    float arrowWidth = 30;
                    float arrowHeight = 30;
                    float posX = winner.Left + (winner.Width / 2) - (arrowWidth / 2);
                    float posY = winner.Top - arrowHeight - 10;
                    
                    if (posY < 0) posY = winner.Top + winner.Height + 10; // Fallback if no space above

                    // MsoAutoShapeType.msoShapeDownArrow
                    var arrow = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeDownArrow, posX, posY, arrowWidth, arrowHeight);
                    arrow.Fill.ForeColor.RGB = 0x0000FF; // Red 
                    arrow.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                    
                    _drawnArrowIds.Add(arrow.Id.ToString());
                }
            }
            catch { }
        }

        private void RemoveArrows()
        {
            try
            {
                var app = _pptManager.GetApplication();
                if (app.ActiveWindow == null || app.ActiveWindow.View.Slide == null) return;
                var slide = (Slide)app.ActiveWindow.View.Slide;

                foreach (var shpIdStr in _drawnArrowIds)
                {
                    try
                    {
                        int shpId = int.Parse(shpIdStr);
                        foreach (Shape shp in slide.Shapes)
                        {
                            if (shp.Id == shpId)
                            {
                                shp.Delete();
                                break;
                            }
                        }
                    }
                    catch { }
                }
                _drawnArrowIds.Clear();
            }
            catch { }
        }
    }
}
