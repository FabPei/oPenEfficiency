using System.Windows;
using oPenEfficiency.Services;
using oPenEfficiency.UI.Panels;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class GanttChartWindow : Window
    {
        private readonly GanttChartPanel _panel;

        public GanttChartWindow(PowerPointManager manager, GanttStartAction startAction = GanttStartAction.Open)
        {
            InitializeComponent();
            _panel = new GanttChartPanel(manager, startAction);
            MainGrid.Children.Add(_panel);
        }

        public void ApplyStartAction(GanttStartAction action) => _panel.ApplyStartAction(action);
    }
}
