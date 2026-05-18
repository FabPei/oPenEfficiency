using oPenEfficiency.UI;
using System.Windows;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Lottery/Random winner picker.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnWinnerPicker",
        Name = "Lottery",
        Tooltip = "Lottery",
        IconData = "M19,3H5C3.9,3 3,3.9 3,5V19C3,20.1 3.9,21 5,21H19C20.1,21 21,20.1 21,19V5C21,3.9 20.1,3 19,3M19,19H5V5H19V19M12,17L14,15H10L12,17M12,7C10.9,7 10,7.9 10,9C10,10.1 10.9,11 12,11C13.1,11 14,10.1 14,9C14,7.9 13.1,7 12,7M7,15H9V13H7V15M15,15H17V13H15V15M7,11H9V9H7V11M15,11H17V9H15V11Z",
        Color = "#F43F5E",
        Description = "Opens the lottery window for random winner selection.",
        DetailedHelpText = "### Lottery\nAnimates a random selection from a defined list of names or items, useful for team engagement exercises and random presenter selection.",
        MinSelection = 0)]
    public static class LotteryFeature
    {
        private static LotteryWindow _lotteryWindow;

        /// <summary>
        /// Standard Execute method for auto-discovery.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return ExecuteWithToggle(manager, null);
        }

        /// <summary>
        /// Specialized Execute method with ToggleButton support for the sidebar.
        /// </summary>
        public static bool ExecuteWithToggle(PowerPointManager manager, System.Windows.Controls.Primitives.ToggleButton toggleBtn)
        {
            if (toggleBtn != null && toggleBtn.IsChecked == false)
            {
                if (_lotteryWindow != null && _lotteryWindow.IsLoaded)
                {
                    _lotteryWindow.Close();
                }
                return true;
            }

            if (_lotteryWindow == null || !_lotteryWindow.IsLoaded)
            {
                _lotteryWindow = new LotteryWindow(manager);
                if (toggleBtn != null)
                {
                    toggleBtn.IsChecked = true;
                    _lotteryWindow.Closed += (s, e) => { toggleBtn.IsChecked = false; };
                }
                _lotteryWindow.Show();
            }
            else
            {
                if (toggleBtn != null) toggleBtn.IsChecked = true;
                _lotteryWindow.Activate();
            }
            return true;
        }
    }
}
