using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;

namespace oPenEfficiency.UI
{
    public abstract class MiniGameWindow : Window
    {
        protected readonly Brush Surface = new SolidColorBrush(Color.FromRgb(30, 30, 30));
        protected readonly Brush Text = new SolidColorBrush(Color.FromRgb(248, 250, 252));
        protected readonly Brush Secondary = new SolidColorBrush(Color.FromRgb(190, 190, 190));

        protected MiniGameWindow(string title, double width, double height)
        {
            Title = title; Width = width; Height = height; MinWidth = width; MinHeight = height;
            WindowStyle = WindowStyle.None; AllowsTransparency = true; Background = Brushes.Transparent; Topmost = true; ShowInTaskbar = false; WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        protected Border CreateShell(string title, out Grid layout)
        {
            var root = new Border { Background = Surface, BorderBrush = new SolidColorBrush(Color.FromRgb(75, 75, 75)), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(12), Effect = new DropShadowEffect { BlurRadius = 20, ShadowDepth = 0, Opacity = .6, Color = Colors.Black } };
            layout = new Grid { Margin = new Thickness(18) }; root.Child = layout;
            var header = new Grid { Margin = new Thickness(0, 0, 0, 10) };
            header.ColumnDefinitions.Add(new ColumnDefinition()); header.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            header.Children.Add(new TextBlock { Text = title.ToUpperInvariant(), FontSize = 20, FontWeight = FontWeights.SemiBold, Foreground = Text, VerticalAlignment = VerticalAlignment.Center });
            var close = Button("✕", new SolidColorBrush(Color.FromRgb(55, 55, 55))); close.Width = 32; close.Height = 32; close.FontSize = 16; close.Padding = new Thickness(0); close.HorizontalContentAlignment = HorizontalAlignment.Center; close.VerticalContentAlignment = VerticalAlignment.Center; close.Click += (s, e) => Close();
            Grid.SetColumn(close, 1); header.Children.Add(close); layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); Grid.SetRow(header, 0); layout.Children.Add(header);
            header.MouseDown += (s, e) => { if (e.ChangedButton == MouseButton.Left) DragMove(); };
            return root;
        }

        protected Button Button(string label, Brush background)
        {
            return new Button { Content = label, Background = background, Foreground = Text, BorderThickness = new Thickness(0), Padding = new Thickness(12, 6, 12, 6), Cursor = Cursors.Hand };
        }
    }

    public sealed class Game2048Window : MiniGameWindow
    {
        private readonly int[,] _board = new int[4, 4];
        private readonly Border[,] _tiles = new Border[4, 4];
        private readonly TextBlock _score = new TextBlock();
        private readonly TextBlock _status = new TextBlock();
        private int _points;
        private bool _gameOver;
        private readonly Random _random = new Random();

        public Game2048Window() : base("2048", 500, 665)
        {
            Content = Build(); StartNewGame(); Focusable = true; Loaded += (s, e) => Focus(); PreviewKeyDown += OnKeyDown;
        }

        private UIElement Build()
        {
            Grid layout; var root = CreateShell("2048", out layout);
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); layout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var info = new Grid { Margin = new Thickness(0, 0, 0, 10) }; info.ColumnDefinitions.Add(new ColumnDefinition()); info.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            _status.Foreground = Secondary; _status.FontSize = 13; _status.Text = "Use arrow keys to combine tiles."; info.Children.Add(_status);
            _score.Foreground = Text; _score.FontWeight = FontWeights.Bold; Grid.SetColumn(_score, 1); info.Children.Add(_score); Grid.SetRow(info, 1); layout.Children.Add(info);
            var board = new Grid { Background = new SolidColorBrush(Color.FromRgb(72, 64, 56)), Margin = new Thickness(0, 0, 0, 12) };
            for (var i = 0; i < 4; i++) { board.RowDefinitions.Add(new RowDefinition()); board.ColumnDefinitions.Add(new ColumnDefinition()); }
            for (var row = 0; row < 4; row++) for (var column = 0; column < 4; column++)
            {
                var tileText = new TextBlock { HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, FontWeight = FontWeights.Bold };
                var tile = new Border { Margin = new Thickness(5), CornerRadius = new CornerRadius(4), Child = tileText };
                Grid.SetRow(tile, row); Grid.SetColumn(tile, column); board.Children.Add(tile); _tiles[row, column] = tile;
            }
            Grid.SetRow(board, 2); layout.Children.Add(board);
            var controls = new StackPanel { HorizontalAlignment = HorizontalAlignment.Center };
            var up = Button("↑", new SolidColorBrush(Color.FromRgb(217, 119, 6))); up.Width = 52; up.FontSize = 22; up.Click += (s, e) => ProcessMove(-1, 0); controls.Children.Add(up);
            var horizontal = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 4, 0, 8) };
            var left = Button("←", new SolidColorBrush(Color.FromRgb(217, 119, 6))); left.Width = 52; left.FontSize = 22; left.Click += (s, e) => ProcessMove(0, -1); horizontal.Children.Add(left);
            var down = Button("↓", new SolidColorBrush(Color.FromRgb(217, 119, 6))); down.Width = 52; down.FontSize = 22; down.Margin = new Thickness(4, 0, 4, 0); down.Click += (s, e) => ProcessMove(1, 0); horizontal.Children.Add(down);
            var right = Button("→", new SolidColorBrush(Color.FromRgb(217, 119, 6))); right.Width = 52; right.FontSize = 22; right.Click += (s, e) => ProcessMove(0, 1); horizontal.Children.Add(right); controls.Children.Add(horizontal);
            var reset = Button("New game", new SolidColorBrush(Color.FromRgb(95, 85, 75))); reset.HorizontalAlignment = HorizontalAlignment.Center; reset.Click += (s, e) => StartNewGame(); controls.Children.Add(reset); Grid.SetRow(controls, 3); layout.Children.Add(controls);
            return root;
        }

        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Left) ProcessMove(0, -1); else if (e.Key == Key.Right) ProcessMove(0, 1); else if (e.Key == Key.Up) ProcessMove(-1, 0); else if (e.Key == Key.Down) ProcessMove(1, 0); else return;
            e.Handled = true;
        }
        private void ProcessMove(int rowDirection, int columnDirection)
        {
            if (_gameOver) return;
            if (Move(rowDirection, columnDirection))
            {
                AddTile(); Render();
            }
            if (IsGameOver())
            {
                _gameOver = true;
                _status.Text = "Game over — no moves left.";
                MessageBox.Show("No moves left. Your score: " + _points, "2048 — Game over", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        private bool Move(int rowDirection, int columnDirection)
        {
            var moved = false;
            for (var line = 0; line < 4; line++)
            {
                var values = new List<int>();
                for (var position = 0; position < 4; position++)
                {
                    var row = rowDirection == 0 ? line : (rowDirection < 0 ? position : 3 - position);
                    var column = columnDirection == 0 ? line : (columnDirection < 0 ? position : 3 - position);
                    if (_board[row, column] != 0) values.Add(_board[row, column]);
                }
                var merged = new List<int>();
                for (var index = 0; index < values.Count; index++)
                {
                    if (index + 1 < values.Count && values[index] == values[index + 1]) { var combined = values[index] * 2; merged.Add(combined); _points += combined; index++; }
                    else merged.Add(values[index]);
                }
                while (merged.Count < 4) merged.Add(0);
                for (var position = 0; position < 4; position++)
                {
                    var row = rowDirection == 0 ? line : (rowDirection < 0 ? position : 3 - position);
                    var column = columnDirection == 0 ? line : (columnDirection < 0 ? position : 3 - position);
                    if (_board[row, column] != merged[position]) moved = true;
                    _board[row, column] = merged[position];
                }
            }
            return moved;
        }
        public void StartNewGame() { Array.Clear(_board, 0, _board.Length); _points = 0; _gameOver = false; AddTile(); AddTile(); Render(); Focus(); }
        private void AddTile() { var empty = new List<BoardPosition>(); for (var r = 0; r < 4; r++) for (var c = 0; c < 4; c++) if (_board[r, c] == 0) empty.Add(new BoardPosition(r, c)); if (empty.Count > 0) { var cell = empty[_random.Next(empty.Count)]; _board[cell.Row, cell.Column] = _random.Next(10) == 0 ? 4 : 2; } }
        private void Render()
        {
            var won = _board.Cast<int>().Any(value => value >= 2048); var blocked = !_board.Cast<int>().Any(value => value == 0) && !CanMove(); _status.Text = won ? "You reached 2048!" : blocked ? "No moves left — start a new game." : "Use arrow keys to combine tiles."; _score.Text = "Score: " + _points;
            for (var r = 0; r < 4; r++) for (var c = 0; c < 4; c++) { var value = _board[r, c]; var text = (TextBlock)_tiles[r, c].Child; text.Text = value == 0 ? string.Empty : value.ToString(); text.FontSize = value >= 1024 ? 22 : value >= 128 ? 28 : 34; text.Foreground = value <= 4 ? new SolidColorBrush(Color.FromRgb(70, 65, 60)) : Brushes.White; _tiles[r, c].Background = TileColor(value); }
        }
        private bool CanMove() { for (var r = 0; r < 4; r++) for (var c = 0; c < 4; c++) if ((Inside(r + 1, c) && _board[r, c] == _board[r + 1, c]) || (Inside(r, c + 1) && _board[r, c] == _board[r, c + 1])) return true; return false; }
        private bool IsGameOver() { return !_board.Cast<int>().Any(value => value == 0) && !CanMove(); }
        private static bool Inside(int r, int c) { return r >= 0 && r < 4 && c >= 0 && c < 4; }
        private static Brush TileColor(int value) { var colors = new Dictionary<int, Color> { { 0, Color.FromRgb(110, 101, 92) }, { 2, Color.FromRgb(238, 228, 218) }, { 4, Color.FromRgb(237, 224, 200) }, { 8, Color.FromRgb(242, 177, 121) }, { 16, Color.FromRgb(245, 149, 99) }, { 32, Color.FromRgb(246, 124, 95) }, { 64, Color.FromRgb(246, 94, 59) }, { 128, Color.FromRgb(237, 207, 114) }, { 256, Color.FromRgb(237, 204, 97) }, { 512, Color.FromRgb(237, 200, 80) }, { 1024, Color.FromRgb(237, 197, 63) }, { 2048, Color.FromRgb(237, 194, 46) } }; return new SolidColorBrush(colors.ContainsKey(value) ? colors[value] : Color.FromRgb(60, 58, 50)); }
    }

    public sealed class MinesweeperWindow : MiniGameWindow
    {
        private const int Size = 9, MineCount = 10;
        private readonly MineCell[,] _cells = new MineCell[Size, Size];
        private readonly Button[,] _buttons = new Button[Size, Size];
        private readonly Random _random = new Random();
        private readonly TextBlock _status = new TextBlock();
        private bool _finished;

        public MinesweeperWindow() : base("Minesweeper", 560, 650) { Content = Build(); StartNewGame(); }
        private UIElement Build()
        {
            Grid layout; var root = CreateShell("Minesweeper", out layout);
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); layout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            _status.Foreground = Secondary; _status.FontSize = 13; _status.Margin = new Thickness(0, 0, 0, 10); Grid.SetRow(_status, 1); layout.Children.Add(_status);
            var board = new Grid { Background = new SolidColorBrush(Color.FromRgb(55, 65, 81)), Margin = new Thickness(0, 0, 0, 12) };
            for (var i = 0; i < Size; i++) { board.RowDefinitions.Add(new RowDefinition()); board.ColumnDefinitions.Add(new ColumnDefinition()); }
            for (var row = 0; row < Size; row++) for (var column = 0; column < Size; column++)
            {
                var button = new Button { Margin = new Thickness(1), Padding = new Thickness(0), FontWeight = FontWeights.Bold, FontSize = 15, Tag = new BoardPosition(row, column) }; button.Click += OpenCell; button.PreviewMouseRightButtonDown += MarkCell;
                Grid.SetRow(button, row); Grid.SetColumn(button, column); board.Children.Add(button); _buttons[row, column] = button;
            }
            Grid.SetRow(board, 2); layout.Children.Add(board);
            var reset = Button("New game", new SolidColorBrush(Color.FromRgb(16, 185, 129))); reset.HorizontalAlignment = HorizontalAlignment.Center; reset.Click += (s, e) => StartNewGame(); Grid.SetRow(reset, 3); layout.Children.Add(reset); return root;
        }
        public void StartNewGame()
        {
            _finished = false; for (var r = 0; r < Size; r++) for (var c = 0; c < Size; c++) _cells[r, c] = new MineCell();
            var placed = 0; while (placed < MineCount) { var r = _random.Next(Size); var c = _random.Next(Size); if (_cells[r, c].Mine) continue; _cells[r, c].Mine = true; placed++; }
            for (var r = 0; r < Size; r++) for (var c = 0; c < Size; c++) _cells[r, c].Nearby = NearbyMines(r, c);
            _status.Text = "Clear the field. Right-click to mark a mine."; Render();
        }
        private void OpenCell(object sender, RoutedEventArgs e) { if (_finished) return; var p = (BoardPosition)((Button)sender).Tag; if (_cells[p.Row, p.Column].Marked) return; Reveal(p.Row, p.Column); if (_cells[p.Row, p.Column].Mine) { _finished = true; _status.Text = "Boom — new game?"; } else if (IsWon()) { _finished = true; _status.Text = "Field cleared — you win!"; } Render(); }
        private void MarkCell(object sender, MouseButtonEventArgs e) { if (_finished) return; var p = (BoardPosition)((Button)sender).Tag; var cell = _cells[p.Row, p.Column]; if (!cell.Open) cell.Marked = !cell.Marked; Render(); e.Handled = true; }
        private void Reveal(int row, int column) { if (!Inside(row, column) || _cells[row, column].Open || _cells[row, column].Marked) return; _cells[row, column].Open = true; if (_cells[row, column].Nearby == 0 && !_cells[row, column].Mine) for (var dr = -1; dr <= 1; dr++) for (var dc = -1; dc <= 1; dc++) if (dr != 0 || dc != 0) Reveal(row + dr, column + dc); }
        private void Render()
        {
            for (var r = 0; r < Size; r++) for (var c = 0; c < Size; c++)
            {
                var cell = _cells[r, c]; var button = _buttons[r, c];
                button.IsEnabled = !_finished; button.IsHitTestVisible = !_finished && !cell.Open; button.Opacity = 1;
                button.Background = new SolidColorBrush(cell.Open ? Color.FromRgb(255, 255, 255) : Color.FromRgb(148, 163, 184));
                var label = cell.Open ? (cell.Mine ? "✹" : cell.Nearby == 0 ? string.Empty : cell.Nearby.ToString()) : cell.Marked ? "⚑" : string.Empty;
                var color = cell.Mine ? Brushes.Red : cell.Marked ? new SolidColorBrush(Color.FromRgb(127, 29, 29)) : MineColor(cell.Nearby);
                button.Content = string.IsNullOrEmpty(label)
                    ? null
                    : new TextBlock
                    {
                        Text = label,
                        FontSize = 18,
                        FontWeight = FontWeights.Bold,
                        Foreground = color,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center
                    };
            }
        }
        private int NearbyMines(int row, int column) { var count = 0; for (var dr = -1; dr <= 1; dr++) for (var dc = -1; dc <= 1; dc++) if ((dr != 0 || dc != 0) && Inside(row + dr, column + dc) && _cells[row + dr, column + dc].Mine) count++; return count; }
        private bool IsWon() { return _cells.Cast<MineCell>().All(cell => cell.Mine || cell.Open); }
        private static bool Inside(int r, int c) { return r >= 0 && r < Size && c >= 0 && c < Size; }
        private static Brush MineColor(int nearby) { return new SolidColorBrush(nearby == 1 ? Color.FromRgb(37, 99, 235) : nearby == 2 ? Color.FromRgb(22, 163, 74) : nearby == 3 ? Color.FromRgb(220, 38, 38) : Color.FromRgb(79, 70, 229)); }
        private sealed class MineCell { public bool Mine, Open, Marked; public int Nearby; }
    }
}
