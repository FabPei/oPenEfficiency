using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    /// <summary>A local, presentation-independent chess game. White is controlled by the user.</summary>
    public sealed class ChessWindow : Window
    {
        private const int BoardSize = 8;
        private readonly Button[,] _squares = new Button[BoardSize, BoardSize];
        private readonly ChessGame _game = new ChessGame();
        private readonly TextBlock _status = new TextBlock();
        private readonly ListBox _moveHistory = new ListBox();
        private readonly ComboBox _difficulty = new ComboBox();
        private int _selectedRow = -1;
        private int _selectedColumn = -1;
        private int _computerDepth = 2;
        private bool _computerThinking;

        public ChessWindow()
        {
            Title = "Chess";
            Width = 760;
            Height = 690;
            MinWidth = 680;
            MinHeight = 610;
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            Topmost = true;
            ShowInTaskbar = false;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;

            Content = CreateContent();
            RenderBoard();
        }

        private UIElement CreateContent()
        {
            var root = new Border
            {
                Background = FindBrush("SurfaceBrush", Color.FromRgb(30, 30, 30)),
                BorderBrush = FindBrush("BorderLightBrush", Color.FromRgb(75, 75, 75)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(12),
                Effect = new DropShadowEffect { BlurRadius = 20, ShadowDepth = 0, Opacity = 0.6, Color = Colors.Black }
            };
            root.MouseDown += OnWindowMouseDown;

            var layout = new Grid { Margin = new Thickness(18) };
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            layout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            layout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.Child = layout;

            var header = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            header.MouseDown += OnHeaderMouseDown;
            header.ColumnDefinitions.Add(new ColumnDefinition());
            header.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            var title = new TextBlock { Text = "CHESS", FontSize = 20, FontWeight = FontWeights.SemiBold, Foreground = FindBrush("TextPrimaryBrush", Colors.White), VerticalAlignment = VerticalAlignment.Center };
            header.Children.Add(title);
            var close = CreateButton("×", FindBrush("SurfaceLightBrush", Color.FromRgb(55, 55, 55)));
            close.Width = 32; close.Height = 32;
            close.Padding = new Thickness(0);
            close.HorizontalContentAlignment = HorizontalAlignment.Center;
            close.VerticalContentAlignment = VerticalAlignment.Center;
            close.Content = new TextBlock
            {
                Text = "✕",
                FontSize = 16,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromRgb(248, 250, 252)),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            close.ToolTip = "Close chess";
            close.Click += (s, e) => Close();
            Grid.SetColumn(close, 1); header.Children.Add(close);
            Grid.SetRow(header, 0); layout.Children.Add(header);

            _status.Foreground = FindBrush("TextSecondaryBrush", Color.FromRgb(190, 190, 190));
            _status.FontSize = 13;
            _status.Margin = new Thickness(0, 0, 0, 12);
            Grid.SetRow(_status, 1); layout.Children.Add(_status);

            var board = new Grid { HorizontalAlignment = HorizontalAlignment.Stretch, VerticalAlignment = VerticalAlignment.Stretch };
            for (var i = 0; i < BoardSize; i++) { board.RowDefinitions.Add(new RowDefinition()); board.ColumnDefinitions.Add(new ColumnDefinition()); }
            for (var row = 0; row < BoardSize; row++)
            for (var col = 0; col < BoardSize; col++)
            {
                var square = new Button { Padding = new Thickness(0), BorderThickness = new Thickness(0), FontFamily = new FontFamily("Segoe UI Symbol"), FontSize = 34, Tag = new BoardPosition(row, col) };
                square.Click += OnSquareClick;
                Grid.SetRow(square, row); Grid.SetColumn(square, col);
                board.Children.Add(square); _squares[row, col] = square;
            }
            var gameArea = new Grid();
            gameArea.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            gameArea.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(170) });
            Grid.SetColumn(board, 0); gameArea.Children.Add(board);

            var sidePanel = new StackPanel { Margin = new Thickness(14, 0, 0, 0) };
            sidePanel.Children.Add(new TextBlock { Text = "CPU STRENGTH", FontSize = 11, FontWeight = FontWeights.Bold, Foreground = FindBrush("TextSecondaryBrush", Color.FromRgb(190, 190, 190)), Margin = new Thickness(0, 0, 0, 5) });
            _difficulty.SelectedIndex = 1;
            _difficulty.Margin = new Thickness(0, 0, 0, 14);
            _difficulty.Items.Add(new ComboBoxItem { Content = "Easy", Tag = 1 });
            _difficulty.Items.Add(new ComboBoxItem { Content = "Medium", Tag = 2 });
            _difficulty.Items.Add(new ComboBoxItem { Content = "Hard", Tag = 3 });
            _difficulty.SelectionChanged += (s, e) =>
            {
                var item = _difficulty.SelectedItem as ComboBoxItem;
                if (item != null) _computerDepth = (int)item.Tag;
            };
            sidePanel.Children.Add(_difficulty);
            sidePanel.Children.Add(new TextBlock { Text = "MOVE HISTORY", FontSize = 11, FontWeight = FontWeights.Bold, Foreground = FindBrush("TextSecondaryBrush", Color.FromRgb(190, 190, 190)), Margin = new Thickness(0, 0, 0, 5) });
            _moveHistory.Background = FindBrush("SurfaceLightBrush", Color.FromRgb(55, 55, 55));
            _moveHistory.Foreground = FindBrush("TextPrimaryBrush", Colors.White);
            _moveHistory.BorderBrush = FindBrush("BorderLightBrush", Color.FromRgb(75, 75, 75));
            _moveHistory.BorderThickness = new Thickness(1);
            _moveHistory.FontSize = 13;
            _moveHistory.Height = 400;
            _moveHistory.VerticalContentAlignment = VerticalAlignment.Top;
            sidePanel.Children.Add(_moveHistory);
            Grid.SetColumn(sidePanel, 1); gameArea.Children.Add(sidePanel);

            Grid.SetRow(gameArea, 2); layout.Children.Add(gameArea);

            var actions = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 14, 0, 0) };
            var restart = CreateButton("New game", FindBrush("AccentBrush", Color.FromRgb(99, 102, 241)));
            restart.Click += (s, e) => { _game.Reset(); ClearSelection(); RenderBoard(); };
            actions.Children.Add(restart);
            var undo = CreateButton("Undo move", FindBrush("SurfaceLightBrush", Color.FromRgb(55, 55, 55)));
            undo.Margin = new Thickness(8, 0, 0, 0); undo.Click += (s, e) => { if (!_computerThinking && _game.UndoTurn()) { ClearSelection(); RenderBoard(); } };
            actions.Children.Add(undo);
            Grid.SetRow(actions, 3); layout.Children.Add(actions);

            return root;
        }

        public void StartNewGame(int computerDepth)
        {
            _computerDepth = Math.Max(1, Math.Min(3, computerDepth));
            _difficulty.SelectedIndex = _computerDepth - 1;
            _game.Reset();
            ClearSelection();
            RenderBoard();
            Activate();
        }

        private Button CreateButton(string text, Brush background)
        {
            return new Button { Content = text, Background = background, Foreground = FindBrush("TextPrimaryBrush", Colors.White), BorderThickness = new Thickness(0), Padding = new Thickness(14, 7, 14, 7), Cursor = Cursors.Hand };
        }

        private void OnWindowMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && e.OriginalSource == sender) DragMove();
        }

        private void OnHeaderMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void OnSquareClick(object sender, RoutedEventArgs e)
        {
            if (_computerThinking || _game.IsFinished || _game.WhiteToMove == false) return;
            var position = (BoardPosition)((Button)sender).Tag;
            var piece = _game[position.Row, position.Column];
            if (_selectedRow >= 0 && _game.IsLegalMove(_selectedRow, _selectedColumn, position.Row, position.Column))
            {
                var move = new ChessMove(_selectedRow, _selectedColumn, position.Row, position.Column);
                if (_game.RequiresPromotion(move)) move.Promotion = PromotionWindow.Choose(this);
                _game.MakeMove(move);
                ClearSelection(); RenderBoard();
                if (!_game.IsFinished) MakeComputerMove();
                return;
            }
            if (piece != null && piece.IsWhite) { _selectedRow = position.Row; _selectedColumn = position.Column; RenderBoard(); }
            else ClearSelection();
        }

        private void MakeComputerMove()
        {
            _computerThinking = true;
            _status.Text = "Computer is thinking…";
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try { _game.MakeBestMove(_computerDepth); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "ChessWindow.MakeComputerMove"); }
                finally { _computerThinking = false; RenderBoard(); }
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        private void ClearSelection() { _selectedRow = -1; _selectedColumn = -1; }

        private void RenderBoard()
        {
            var legalTargets = _selectedRow < 0 ? new List<ChessMove>() : _game.GetLegalMoves().Where(m => m.FromRow == _selectedRow && m.FromColumn == _selectedColumn).ToList();
            for (var row = 0; row < BoardSize; row++)
            for (var col = 0; col < BoardSize; col++)
            {
                var target = legalTargets.Any(m => m.ToRow == row && m.ToColumn == col);
                var selected = row == _selectedRow && col == _selectedColumn;
                var light = (row + col) % 2 == 0;
                _squares[row, col].Background = new SolidColorBrush(selected ? Color.FromRgb(234, 179, 8) : target ? Color.FromRgb(134, 174, 104) : light ? Color.FromRgb(240, 217, 181) : Color.FromRgb(181, 136, 99));
                var piece = _game[row, col];
                _squares[row, col].Content = piece == null
                    ? null
                    : new TextBlock
                    {
                        Text = piece.Symbol,
                        FontFamily = new FontFamily("Segoe UI Symbol"),
                        FontSize = 34,
                        FontWeight = FontWeights.SemiBold,
                        Foreground = piece.IsWhite
                            ? new SolidColorBrush(Color.FromRgb(248, 250, 252))
                            : new SolidColorBrush(Color.FromRgb(17, 24, 39)),
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center
                    };
            }
            _moveHistory.Items.Clear();
            for (var moveIndex = 0; moveIndex < _game.MoveHistory.Count; moveIndex += 2)
            {
                var number = moveIndex / 2 + 1;
                var blackMove = moveIndex + 1 < _game.MoveHistory.Count ? "  " + _game.MoveHistory[moveIndex + 1] : string.Empty;
                _moveHistory.Items.Add(number + ". " + _game.MoveHistory[moveIndex] + blackMove);
            }
            if (_moveHistory.Items.Count > 0) _moveHistory.ScrollIntoView(_moveHistory.Items[_moveHistory.Items.Count - 1]);
            _status.Text = _game.Status;
        }

        private Brush FindBrush(string key, Color fallback)
        {
            return Application.Current != null && Application.Current.Resources[key] is Brush brush ? brush : new SolidColorBrush(fallback);
        }
    }

    internal sealed class PromotionWindow : Window
    {
        private PieceType _choice = PieceType.Queen;

        private PromotionWindow(Window owner)
        {
            Owner = owner; Width = 310; Height = 155; WindowStyle = WindowStyle.None; AllowsTransparency = true; Background = Brushes.Transparent; ShowInTaskbar = false; WindowStartupLocation = WindowStartupLocation.CenterOwner;
            var shell = new Border { Background = new SolidColorBrush(Color.FromRgb(30, 30, 30)), BorderBrush = new SolidColorBrush(Color.FromRgb(75, 75, 75)), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(10), Padding = new Thickness(16), Effect = new DropShadowEffect { BlurRadius = 16, ShadowDepth = 0, Opacity = .6, Color = Colors.Black } };
            var panel = new StackPanel(); shell.Child = panel; panel.Children.Add(new TextBlock { Text = "Promote pawn to", Foreground = Brushes.White, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 12) });
            var choices = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center };
            foreach (var type in new[] { PieceType.Queen, PieceType.Rook, PieceType.Bishop, PieceType.Knight })
            {
                var selected = type;
                var button = new Button { Content = "♕♖♗♘".Substring(type == PieceType.Queen ? 0 : type == PieceType.Rook ? 1 : type == PieceType.Bishop ? 2 : 3, 1), FontFamily = new FontFamily("Segoe UI Symbol"), FontSize = 25, Foreground = Brushes.White, Background = new SolidColorBrush(Color.FromRgb(75, 65, 55)), BorderThickness = new Thickness(0), Width = 52, Height = 45, Margin = new Thickness(3, 0, 3, 0), Cursor = Cursors.Hand };
                button.Click += (s, e) => { _choice = selected; DialogResult = true; };
                choices.Children.Add(button);
            }
            panel.Children.Add(choices); Content = shell;
        }

        public static PieceType Choose(Window owner) { var dialog = new PromotionWindow(owner); dialog.ShowDialog(); return dialog._choice; }
    }

    internal struct BoardPosition { public int Row; public int Column; public BoardPosition(int row, int column) { Row = row; Column = column; } }
    internal enum PieceType { Pawn, Knight, Bishop, Rook, Queen, King }
    internal sealed class ChessPiece
    {
        public PieceType Type; public bool IsWhite;
        public ChessPiece(PieceType type, bool isWhite) { Type = type; IsWhite = isWhite; }
        public string Symbol { get { return IsWhite ? "♙♘♗♖♕♔".Substring((int)Type, 1) : "♟♞♝♜♛♚".Substring((int)Type, 1); } }
        public ChessPiece Copy() { return new ChessPiece(Type, IsWhite); }
    }
    internal sealed class ChessMove
    {
        public int FromRow, FromColumn, ToRow, ToColumn;
        public PieceType? Promotion;
        public ChessMove(int fr, int fc, int tr, int tc) { FromRow = fr; FromColumn = fc; ToRow = tr; ToColumn = tc; }
    }

    internal sealed class ChessGame
    {
        private ChessPiece[,] _board = new ChessPiece[8, 8];
        private readonly Stack<GameState> _history = new Stack<GameState>();
        private readonly List<string> _moveHistory = new List<string>();
        private readonly List<string> _positionHistory = new List<string>();
        private bool _whiteKingMoved, _blackKingMoved, _whiteLeftRookMoved, _whiteRightRookMoved, _blackLeftRookMoved, _blackRightRookMoved;
        private BoardPosition? _enPassant;
        private int _halfMoveClock;
        public bool WhiteToMove { get; private set; }
        public bool IsFinished { get; private set; }
        public string Status { get; private set; }
        public IList<string> MoveHistory { get { return _moveHistory.AsReadOnly(); } }
        public ChessPiece this[int row, int col] { get { return _board[row, col]; } }

        public ChessGame() { Reset(); }
        public void Reset()
        {
            _board = new ChessPiece[8, 8]; _history.Clear(); _moveHistory.Clear(); _positionHistory.Clear(); _halfMoveClock = 0; WhiteToMove = true; IsFinished = false; Status = "Your turn (White).";
            _whiteKingMoved = _blackKingMoved = _whiteLeftRookMoved = _whiteRightRookMoved = _blackLeftRookMoved = _blackRightRookMoved = false; _enPassant = null;
            var order = new[] { PieceType.Rook, PieceType.Knight, PieceType.Bishop, PieceType.Queen, PieceType.King, PieceType.Bishop, PieceType.Knight, PieceType.Rook };
            for (var c = 0; c < 8; c++) { _board[0, c] = new ChessPiece(order[c], false); _board[1, c] = new ChessPiece(PieceType.Pawn, false); _board[6, c] = new ChessPiece(PieceType.Pawn, true); _board[7, c] = new ChessPiece(order[c], true); }
            _positionHistory.Add(GetPositionKey());
        }
        public bool IsLegalMove(int fr, int fc, int tr, int tc) { return GetLegalMoves().Any(m => m.FromRow == fr && m.FromColumn == fc && m.ToRow == tr && m.ToColumn == tc); }
        public bool RequiresPromotion(ChessMove move) { var piece = _board[move.FromRow, move.FromColumn]; return piece != null && piece.Type == PieceType.Pawn && (move.ToRow == 0 || move.ToRow == 7); }
        public List<ChessMove> GetLegalMoves()
        {
            var moves = GetPseudoLegalMoves(); var legal = new List<ChessMove>();
            foreach (var move in moves) { var state = Capture(); Apply(move); var ownInCheck = IsKingInCheck(!WhiteToMove); Restore(state); if (!ownInCheck) legal.Add(move); }
            return legal;
        }
        public void MakeMove(ChessMove move)
        {
            _history.Push(Capture());
            var notation = GetMoveNotation(move);
            Apply(move);
            _positionHistory.Add(GetPositionKey());
            UpdateStatus();
            if (IsFinished && Status.IndexOf("Checkmate", StringComparison.OrdinalIgnoreCase) >= 0) notation += "#";
            else if (IsKingInCheck(WhiteToMove)) notation += "+";
            _moveHistory.Add(notation);
        }
        public bool UndoTurn()
        {
            if (_history.Count < 2) return false;
            Restore(_history.Pop()); Restore(_history.Pop());
            _moveHistory.RemoveRange(Math.Max(0, _moveHistory.Count - 2), Math.Min(2, _moveHistory.Count));
            _positionHistory.RemoveRange(Math.Max(1, _positionHistory.Count - 2), Math.Min(2, Math.Max(0, _positionHistory.Count - 1)));
            IsFinished = false; Status = "Your turn (White)."; return true;
        }
        public void MakeBestMove(int depth)
        {
            var moves = GetLegalMoves(); if (moves.Count == 0) { UpdateStatus(); return; }
            ChessMove best = moves[0]; var bestScore = int.MaxValue;
            foreach (var move in moves) { var state = Capture(); Apply(move); var score = Search(depth - 1, -100000, 100000); Restore(state); if (score < bestScore) { bestScore = score; best = move; } }
            MakeMove(best);
        }
        private int Search(int depth, int alpha, int beta)
        {
            var moves = GetLegalMoves(); if (depth == 0 || moves.Count == 0) return Evaluate(moves.Count == 0);
            if (WhiteToMove) { var value = -100000; foreach (var m in moves) { var state = Capture(); Apply(m); value = Math.Max(value, Search(depth - 1, alpha, beta)); Restore(state); alpha = Math.Max(alpha, value); if (alpha >= beta) break; } return value; }
            var min = 100000; foreach (var m in moves) { var state = Capture(); Apply(m); min = Math.Min(min, Search(depth - 1, alpha, beta)); Restore(state); beta = Math.Min(beta, min); if (alpha >= beta) break; } return min;
        }
        private int Evaluate(bool noMoves)
        {
            if (noMoves) return IsKingInCheck(WhiteToMove) ? (WhiteToMove ? -90000 : 90000) : 0;
            var values = new[] { 100, 320, 330, 500, 900, 20000 }; var score = 0;
            for (var r = 0; r < 8; r++) for (var c = 0; c < 8; c++) if (_board[r, c] != null) score += (_board[r, c].IsWhite ? 1 : -1) * values[(int)_board[r, c].Type];
            return score;
        }
        private List<ChessMove> GetPseudoLegalMoves()
        {
            var result = new List<ChessMove>();
            for (var r = 0; r < 8; r++) for (var c = 0; c < 8; c++) { var p = _board[r, c]; if (p != null && p.IsWhite == WhiteToMove) AddMoves(result, r, c, p); }
            return result;
        }
        private void AddMoves(List<ChessMove> moves, int r, int c, ChessPiece p)
        {
            int direction = p.IsWhite ? -1 : 1;
            if (p.Type == PieceType.Pawn)
            {
                if (Inside(r + direction, c) && _board[r + direction, c] == null) { AddPawnMove(moves, r, c, r + direction, c); if (r == (p.IsWhite ? 6 : 1) && _board[r + 2 * direction, c] == null) AddPawnMove(moves, r, c, r + 2 * direction, c); }
                foreach (var dc in new[] { -1, 1 }) if (Inside(r + direction, c + dc) && ((_board[r + direction, c + dc] != null && _board[r + direction, c + dc].IsWhite != p.IsWhite) || (_enPassant.HasValue && _enPassant.Value.Row == r + direction && _enPassant.Value.Column == c + dc))) AddPawnMove(moves, r, c, r + direction, c + dc);
                return;
            }
            if (p.Type == PieceType.Knight) { foreach (var d in new[] { new BoardPosition(-2,-1),new BoardPosition(-2,1),new BoardPosition(-1,-2),new BoardPosition(-1,2),new BoardPosition(1,-2),new BoardPosition(1,2),new BoardPosition(2,-1),new BoardPosition(2,1) }) AddIfFree(moves, r, c, r + d.Row, c + d.Column, p); return; }
            if (p.Type == PieceType.King) { for (var dr = -1; dr <= 1; dr++) for (var dc = -1; dc <= 1; dc++) if (dr != 0 || dc != 0) AddIfFree(moves, r, c, r + dr, c + dc, p); AddCastling(moves, r, c, p); return; }
            var directions = p.Type == PieceType.Bishop ? new[] { new BoardPosition(-1,-1),new BoardPosition(-1,1),new BoardPosition(1,-1),new BoardPosition(1,1) } : p.Type == PieceType.Rook ? new[] { new BoardPosition(-1,0),new BoardPosition(1,0),new BoardPosition(0,-1),new BoardPosition(0,1) } : new[] { new BoardPosition(-1,-1),new BoardPosition(-1,1),new BoardPosition(1,-1),new BoardPosition(1,1),new BoardPosition(-1,0),new BoardPosition(1,0),new BoardPosition(0,-1),new BoardPosition(0,1) };
            foreach (var d in directions) { var rr = r + d.Row; var cc = c + d.Column; while (Inside(rr, cc)) { if (_board[rr, cc] == null) Add(moves, r, c, rr, cc); else { if (_board[rr, cc].IsWhite != p.IsWhite) Add(moves, r, c, rr, cc); break; } rr += d.Row; cc += d.Column; } }
        }
        private void AddCastling(List<ChessMove> moves, int r, int c, ChessPiece king)
        {
            var movedKing = king.IsWhite ? _whiteKingMoved : _blackKingMoved; if (movedKing || IsKingInCheck(king.IsWhite)) return;
            var leftMoved = king.IsWhite ? _whiteLeftRookMoved : _blackLeftRookMoved; var rightMoved = king.IsWhite ? _whiteRightRookMoved : _blackRightRookMoved;
            if (!rightMoved && _board[r, 5] == null && _board[r, 6] == null && !IsSquareAttacked(r, 5, !king.IsWhite) && !IsSquareAttacked(r, 6, !king.IsWhite)) Add(moves, r, c, r, 6);
            if (!leftMoved && _board[r, 1] == null && _board[r, 2] == null && _board[r, 3] == null && !IsSquareAttacked(r, 3, !king.IsWhite) && !IsSquareAttacked(r, 2, !king.IsWhite)) Add(moves, r, c, r, 2);
        }
        private void AddIfFree(List<ChessMove> moves, int r, int c, int tr, int tc, ChessPiece p) { if (Inside(tr, tc) && (_board[tr, tc] == null || _board[tr, tc].IsWhite != p.IsWhite)) Add(moves, r, c, tr, tc); }
        private static void AddPawnMove(List<ChessMove> moves, int fr, int fc, int tr, int tc)
        {
            if (tr != 0 && tr != 7) { Add(moves, fr, fc, tr, tc); return; }
            foreach (var promotion in new[] { PieceType.Queen, PieceType.Rook, PieceType.Bishop, PieceType.Knight }) moves.Add(new ChessMove(fr, fc, tr, tc) { Promotion = promotion });
        }
        private static void Add(List<ChessMove> moves, int fr, int fc, int tr, int tc) { moves.Add(new ChessMove(fr, fc, tr, tc)); }
        private static bool Inside(int r, int c) { return r >= 0 && r < 8 && c >= 0 && c < 8; }
        private string GetMoveNotation(ChessMove move)
        {
            var piece = _board[move.FromRow, move.FromColumn];
            if (piece.Type == PieceType.King && Math.Abs(move.ToColumn - move.FromColumn) == 2) return move.ToColumn == 6 ? "O-O" : "O-O-O";
            var captured = _board[move.ToRow, move.ToColumn] != null || (piece.Type == PieceType.Pawn && _enPassant.HasValue && _enPassant.Value.Row == move.ToRow && _enPassant.Value.Column == move.ToColumn);
            var name = piece.Type == PieceType.Pawn ? string.Empty : "NBRQK".Substring((int)piece.Type - 1, 1);
            if (piece.Type == PieceType.Pawn && captured) name = ((char)('a' + move.FromColumn)).ToString();
            var target = ((char)('a' + move.ToColumn)).ToString() + (8 - move.ToRow).ToString();
            var promotion = move.Promotion.HasValue ? "=" + "NBRQK".Substring((int)move.Promotion.Value - 1, 1) : string.Empty;
            return name + (captured ? "x" : string.Empty) + target + promotion;
        }
        private void Apply(ChessMove move)
        {
            var piece = _board[move.FromRow, move.FromColumn]; var captured = _board[move.ToRow, move.ToColumn];
            var enPassantCapture = piece.Type == PieceType.Pawn && _enPassant.HasValue && move.ToRow == _enPassant.Value.Row && move.ToColumn == _enPassant.Value.Column && captured == null;
            _halfMoveClock = piece.Type == PieceType.Pawn || captured != null || enPassantCapture ? 0 : _halfMoveClock + 1;
            if (enPassantCapture) _board[move.FromRow, move.ToColumn] = null;
            _board[move.ToRow, move.ToColumn] = piece; _board[move.FromRow, move.FromColumn] = null;
            if (piece.Type == PieceType.King && Math.Abs(move.ToColumn - move.FromColumn) == 2) { var rookFrom = move.ToColumn == 6 ? 7 : 0; var rookTo = move.ToColumn == 6 ? 5 : 3; _board[move.ToRow, rookTo] = _board[move.ToRow, rookFrom]; _board[move.ToRow, rookFrom] = null; }
            UpdateRights(piece, move); if (captured != null && captured.Type == PieceType.Rook) UpdateRookCapture(move.ToRow, move.ToColumn);
            _enPassant = piece.Type == PieceType.Pawn && Math.Abs(move.ToRow - move.FromRow) == 2 ? new BoardPosition((move.FromRow + move.ToRow) / 2, move.FromColumn) : (BoardPosition?)null;
            if (piece.Type == PieceType.Pawn && (move.ToRow == 0 || move.ToRow == 7)) piece.Type = move.Promotion ?? PieceType.Queen;
            WhiteToMove = !WhiteToMove;
        }
        private void UpdateRights(ChessPiece p, ChessMove m) { if (p.Type == PieceType.King) { if (p.IsWhite) _whiteKingMoved = true; else _blackKingMoved = true; } if (p.Type == PieceType.Rook) UpdateRookCapture(m.FromRow, m.FromColumn); }
        private void UpdateRookCapture(int r, int c) { if (r == 7 && c == 0) _whiteLeftRookMoved = true; if (r == 7 && c == 7) _whiteRightRookMoved = true; if (r == 0 && c == 0) _blackLeftRookMoved = true; if (r == 0 && c == 7) _blackRightRookMoved = true; }
        private bool IsKingInCheck(bool white) { for (var r = 0; r < 8; r++) for (var c = 0; c < 8; c++) if (_board[r, c] != null && _board[r, c].Type == PieceType.King && _board[r, c].IsWhite == white) return IsSquareAttacked(r, c, !white); return true; }
        private bool IsSquareAttacked(int row, int col, bool byWhite)
        {
            var pawnDirection = byWhite ? -1 : 1;
            foreach (var dc in new[] { -1, 1 })
            {
                var pawnRow = row - pawnDirection;
                var pawnColumn = col - dc;
                if (Inside(pawnRow, pawnColumn) && _board[pawnRow, pawnColumn] != null && _board[pawnRow, pawnColumn].IsWhite == byWhite && _board[pawnRow, pawnColumn].Type == PieceType.Pawn) return true;
            }
            foreach (var offset in new[] { new BoardPosition(-2,-1),new BoardPosition(-2,1),new BoardPosition(-1,-2),new BoardPosition(-1,2),new BoardPosition(1,-2),new BoardPosition(1,2),new BoardPosition(2,-1),new BoardPosition(2,1) })
            {
                var r = row + offset.Row; var c = col + offset.Column;
                if (Inside(r, c) && _board[r, c] != null && _board[r, c].IsWhite == byWhite && _board[r, c].Type == PieceType.Knight) return true;
            }
            for (var dr = -1; dr <= 1; dr++) for (var dc = -1; dc <= 1; dc++) if ((dr != 0 || dc != 0) && Inside(row + dr, col + dc) && _board[row + dr, col + dc] != null && _board[row + dr, col + dc].IsWhite == byWhite && _board[row + dr, col + dc].Type == PieceType.King) return true;
            return AttackedOnLine(row, col, byWhite, -1, 0, PieceType.Rook, PieceType.Queen)
                || AttackedOnLine(row, col, byWhite, 1, 0, PieceType.Rook, PieceType.Queen)
                || AttackedOnLine(row, col, byWhite, 0, -1, PieceType.Rook, PieceType.Queen)
                || AttackedOnLine(row, col, byWhite, 0, 1, PieceType.Rook, PieceType.Queen)
                || AttackedOnLine(row, col, byWhite, -1, -1, PieceType.Bishop, PieceType.Queen)
                || AttackedOnLine(row, col, byWhite, -1, 1, PieceType.Bishop, PieceType.Queen)
                || AttackedOnLine(row, col, byWhite, 1, -1, PieceType.Bishop, PieceType.Queen)
                || AttackedOnLine(row, col, byWhite, 1, 1, PieceType.Bishop, PieceType.Queen);
        }
        private bool AttackedOnLine(int row, int col, bool byWhite, int dr, int dc, PieceType first, PieceType second)
        {
            row += dr; col += dc;
            while (Inside(row, col))
            {
                var piece = _board[row, col];
                if (piece != null) return piece.IsWhite == byWhite && (piece.Type == first || piece.Type == second);
                row += dr; col += dc;
            }
            return false;
        }
        private void UpdateStatus()
        {
            var moves = GetLegalMoves(); if (moves.Count == 0) { IsFinished = true; Status = IsKingInCheck(WhiteToMove) ? (WhiteToMove ? "Checkmate — computer wins." : "Checkmate — you win!") : "Draw by stalemate."; return; }
            if (_halfMoveClock >= 100) { IsFinished = true; Status = "Draw by the 50-move rule."; return; }
            if (_positionHistory.Count(key => key == GetPositionKey()) >= 3) { IsFinished = true; Status = "Draw by threefold repetition."; return; }
            if (HasInsufficientMaterial()) { IsFinished = true; Status = "Draw by insufficient material."; return; }
            var check = IsKingInCheck(WhiteToMove) ? " Check!" : string.Empty; Status = (WhiteToMove ? "Your turn (White)." : "Computer's turn (Black).") + check;
        }
        private bool HasInsufficientMaterial()
        {
            var pieces = new List<Tuple<ChessPiece, int, int>>();
            for (var r = 0; r < 8; r++) for (var c = 0; c < 8; c++) if (_board[r, c] != null && _board[r, c].Type != PieceType.King) pieces.Add(Tuple.Create(_board[r, c], r, c));
            if (pieces.Count == 0) return true;
            if (pieces.Count == 1) return pieces[0].Item1.Type == PieceType.Bishop || pieces[0].Item1.Type == PieceType.Knight;
            return pieces.All(value => value.Item1.Type == PieceType.Bishop) && pieces.Select(value => (value.Item2 + value.Item3) % 2).Distinct().Count() == 1;
        }
        private string GetPositionKey()
        {
            var key = new StringBuilder();
            for (var r = 0; r < 8; r++) for (var c = 0; c < 8; c++) key.Append(_board[r, c] == null ? '.' : (char)((_board[r, c].IsWhite ? 'A' : 'a') + (int)_board[r, c].Type));
            key.Append(WhiteToMove ? 'w' : 'b').Append(_whiteKingMoved ? '1' : '0').Append(_blackKingMoved ? '1' : '0').Append(_whiteLeftRookMoved ? '1' : '0').Append(_whiteRightRookMoved ? '1' : '0').Append(_blackLeftRookMoved ? '1' : '0').Append(_blackRightRookMoved ? '1' : '0');
            if (_enPassant.HasValue) key.Append(_enPassant.Value.Row).Append(_enPassant.Value.Column);
            return key.ToString();
        }
        private GameState Capture() { return new GameState { Board = CopyBoard(), White = WhiteToMove, Finished = IsFinished, Text = Status, WK = _whiteKingMoved, BK = _blackKingMoved, WL = _whiteLeftRookMoved, WR = _whiteRightRookMoved, BL = _blackLeftRookMoved, BR = _blackRightRookMoved, EnPassant = _enPassant, HalfMoveClock = _halfMoveClock }; }
        private void Restore(GameState s) { _board = s.Board; WhiteToMove = s.White; IsFinished = s.Finished; Status = s.Text; _whiteKingMoved = s.WK; _blackKingMoved = s.BK; _whiteLeftRookMoved = s.WL; _whiteRightRookMoved = s.WR; _blackLeftRookMoved = s.BL; _blackRightRookMoved = s.BR; _enPassant = s.EnPassant; _halfMoveClock = s.HalfMoveClock; }
        private ChessPiece[,] CopyBoard() { var copy = new ChessPiece[8, 8]; for (var r = 0; r < 8; r++) for (var c = 0; c < 8; c++) if (_board[r, c] != null) copy[r, c] = _board[r, c].Copy(); return copy; }
        private sealed class GameState { public ChessPiece[,] Board; public bool White, Finished, WK, BK, WL, WR, BL, BR; public string Text; public BoardPosition? EnPassant; public int HalfMoveClock; }
    }
}
