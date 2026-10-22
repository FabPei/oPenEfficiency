using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI;

namespace oPenEfficiency.Features
{
    /// <summary>Opens a self-contained chess game against the local computer player.</summary>
    [FeatureMetadata(
        Id = "BtnChess",
        Name = "Chess",
        Tooltip = "Play chess against the computer",
        IconData = "M19,22H5V20H19M17,10H7V12H17M18,7H6V9H18M15.5,2H8.5L7,5H17L15.5,2M7,13V18H17V13H7Z",
        Color = "#6366F1",
        Description = "Opens a floating chess board with a local computer opponent.",
        DetailedHelpText = "### Chess\nPlay a complete game of chess against a built-in computer opponent. The game is independent from the current presentation. Right-click starts a new Easy, Medium, or Hard game or closes the board.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class ChessFeature
    {
        private static ChessWindow _window;

        public static bool Execute(PowerPointManager manager)
        {
            if (_window == null || !_window.IsLoaded)
            {
                _window = new ChessWindow();
                _window.Show();
            }
            else
            {
                _window.Activate();
            }

            return true;
        }

        public static bool StartNewGame(PowerPointManager manager, int computerDepth)
        {
            Execute(manager);
            _window?.StartNewGame(computerDepth);
            return true;
        }

        public static bool CloseGame()
        {
            if (_window == null || !_window.IsLoaded) return false;
            _window.Close();
            return true;
        }
    }
}
