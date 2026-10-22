using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "Btn2048",
        Name = "2048",
        Tooltip = "Play 2048",
        IconData = "M4,4H10V10H4V4M14,4H20V10H14V4M4,14H10V20H4V14M14,14H20V20H14V14",
        Color = "#F59E0B",
        Description = "Opens a compact 2048 puzzle game.",
        DetailedHelpText = "### 2048\nUse the arrow keys to combine tiles and reach 2048. Right-click to continue, start a new game, or close the game window.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class Game2048Feature
    {
        private static Game2048Window _window;
        public static bool Execute(PowerPointManager manager)
        {
            if (_window == null || !_window.IsLoaded) { _window = new Game2048Window(); _window.Show(); }
            else _window.Activate();
            return true;
        }

        public static bool StartNewGame(PowerPointManager manager)
        {
            Execute(manager);
            _window?.StartNewGame();
            return true;
        }

        public static bool CloseGame()
        {
            if (_window == null || !_window.IsLoaded) return false;
            _window.Close();
            return true;
        }
    }

    [FeatureMetadata(
        Id = "BtnMinesweeper",
        Name = "Minesweeper",
        Tooltip = "Play Minesweeper",
        IconData = "M12,2A10,10 0 1,0 22,12A10,10 0 0,0 12,2M12,5A7,7 0 1,1 5,12A7,7 0 0,1 12,5M11,7V11H7V13H11V17H13V13H17V11H13V7H11Z",
        Color = "#10B981",
        Description = "Opens a compact Minesweeper game.",
        DetailedHelpText = "### Minesweeper\nClear the field without opening a mine. Right-click cells to mark suspected mines; right-click the Sidebar command to continue, start a new game, or close the game window.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class MinesweeperFeature
    {
        private static MinesweeperWindow _window;
        public static bool Execute(PowerPointManager manager)
        {
            if (_window == null || !_window.IsLoaded) { _window = new MinesweeperWindow(); _window.Show(); }
            else _window.Activate();
            return true;
        }

        public static bool StartNewGame(PowerPointManager manager)
        {
            Execute(manager);
            _window?.StartNewGame();
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
