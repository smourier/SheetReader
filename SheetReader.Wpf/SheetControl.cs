﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using SheetReader.Wpf.Utilities;

namespace SheetReader.Wpf
{
    [TemplatePart(Name = PartScrollViewerName, Type = typeof(ScrollViewer))]
    public class SheetControl : Control
    {
        public const string PartScrollViewerName = "PART_ScrollViewer";

        public static readonly DependencyProperty SheetProperty = DependencyProperty.Register(nameof(Sheet),
            typeof(BookDocumentSheet),
            typeof(SheetControl),
            new UIPropertyMetadata(null, (d, e) => ((SheetControl)d).OnSheetChanged()));

        public static readonly DependencyProperty ColumnWidthProperty = DependencyProperty.Register(nameof(ColumnWidth),
            typeof(double),
            typeof(SheetControl),
            new UIPropertyMetadata(100.0));

        public static readonly DependencyProperty RowMarginProperty = DependencyProperty.Register(nameof(RowMargin),
            typeof(double),
            typeof(SheetControl),
            new UIPropertyMetadata(60.0));

        public static readonly DependencyProperty RowHeightProperty = DependencyProperty.Register(nameof(RowHeight),
            typeof(double),
            typeof(SheetControl),
            new UIPropertyMetadata(20.0));

        public static readonly DependencyProperty LineSizeProperty = DependencyProperty.Register(nameof(LineSize),
            typeof(double),
            typeof(SheetControl),
            new UIPropertyMetadata(0.5));

        public static readonly DependencyProperty HeaderBrushProperty = DependencyProperty.Register(nameof(HeaderBrush),
            typeof(Brush),
            typeof(SheetControl),
            new UIPropertyMetadata(Brushes.LightGray));

        public static readonly DependencyProperty LineBrushProperty = DependencyProperty.Register(nameof(LineBrush),
            typeof(Brush),
            typeof(SheetControl),
            new UIPropertyMetadata(Brushes.Gray));

        public static readonly DependencyProperty SelectionBrushProperty = DependencyProperty.Register(nameof(SelectionBrush),
            typeof(Brush),
            typeof(SheetControl),
            new UIPropertyMetadata(Brushes.Green));

        public static readonly DependencyProperty TextAlignmentProperty = DependencyProperty.Register(nameof(TextAlignment),
            typeof(TextAlignment),
            typeof(SheetControl),
            new UIPropertyMetadata(TextAlignment.Left));

        public static readonly DependencyProperty TextTrimmingProperty = DependencyProperty.Register(nameof(TextTrimming),
            typeof(TextTrimming),
            typeof(SheetControl),
            new UIPropertyMetadata(TextTrimming.CharacterEllipsis));

        public static readonly DependencyProperty CellPaddingProperty = DependencyProperty.Register(nameof(CellPadding),
            typeof(Thickness),
            typeof(SheetControl),
            new UIPropertyMetadata(new Thickness(5, 0, 5, 0)));

        public static readonly RoutedEvent SelectionChangedEvent = EventManager.RegisterRoutedEvent(nameof(SelectionChanged), RoutingStrategy.Bubble, typeof(RoutedEventHandler), typeof(SheetControl));

        public BookDocumentSheet Sheet { get => (BookDocumentSheet)GetValue(SheetProperty); set => SetValue(SheetProperty, value); }
        public double ColumnWidth { get => (double)GetValue(ColumnWidthProperty); set => SetValue(ColumnWidthProperty, value); }
        public double RowMargin { get => (double)GetValue(RowMarginProperty); set => SetValue(RowMarginProperty, value); }
        public double RowHeight { get => (double)GetValue(RowHeightProperty); set => SetValue(RowHeightProperty, value); }
        public double LineSize { get => (double)GetValue(LineSizeProperty); set => SetValue(LineSizeProperty, value); }
        public Brush LineBrush { get => (Brush)GetValue(LineBrushProperty); set => SetValue(LineBrushProperty, value); }
        public Brush HeaderBrush { get => (Brush)GetValue(HeaderBrushProperty); set => SetValue(HeaderBrushProperty, value); }
        public Brush SelectionBrush { get => (Brush)GetValue(SelectionBrushProperty); set => SetValue(SelectionBrushProperty, value); }
        public TextAlignment TextAlignment { get => (TextAlignment)GetValue(TextAlignmentProperty); set => SetValue(TextAlignmentProperty, value); }
        public TextTrimming TextTrimming { get => (TextTrimming)GetValue(TextTrimmingProperty); set => SetValue(TextTrimmingProperty, value); }
        public Thickness CellPadding { get => (Thickness)GetValue(CellPaddingProperty); set => SetValue(CellPaddingProperty, value); }

        public event RoutedEventHandler SelectionChanged { add => AddHandler(SelectionChangedEvent, value); remove => RemoveHandler(SelectionChangedEvent, value); }

        private const double _minWidth = 4;
        private const double _movingColumnTolerance = 4;

        internal double GetRowHeight() => Math.Max(RowHeight, _minWidth);
        internal double GetRowMargin() => Math.Max(RowMargin, _minWidth);
        internal double GetLineSize() => Math.Max(LineSize, 0);
        internal double GetColumnWidth() => Math.Max(ColumnWidth, _minWidth);

        static SheetControl()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(SheetControl), new FrameworkPropertyMetadata(typeof(SheetControl)));
        }

        private static void Log(object? message = null, [CallerMemberName] string? methodName = null) => EventProvider.Default.WriteMessageEvent(methodName + ":" + message);

        private ScrollViewer? _scrollViewer;
        private SheetGrid? _grid;
        private MovingColumn? _movingColumn;
        private bool _extendingSelection;
        private readonly List<SheetControlColumn> _columnSettings = [];

        public SheetControl()
        {
            Selection = new SheetSelection(this);
            Focusable = true;
        }

        public SheetSelection Selection { get; }
        public IReadOnlyList<SheetControlColumn> ColumnSettings => _columnSettings.AsReadOnly();

        public virtual SheetControlHitTestResult HitTest(Point point) => _grid?.HitTest(point) ?? new();
        public virtual void SetColumnSize(int columnIndex, double width)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(width);
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columnIndex, _columnSettings.Count);
            _columnSettings[columnIndex].Width = width;
            _grid?.InvalidateMeasure();
        }

        public bool EnsureVisible(RowCol rowCol) { ArgumentNullException.ThrowIfNull(rowCol); return EnsureVisible(rowCol.RowIndex, rowCol.ColumnIndex); }
        public virtual bool EnsureVisible(int rowIndex, int columnIndex)
        {
            if (_scrollViewer == null)
                return false;

            var selection = SheetSelection.From(this, rowIndex, columnIndex);
            if (selection == null)
                return false;

            var context = CreateStyleContext();
            if (context == null)
                throw new InvalidOperationException();

            context.RowHeight ??= GetRowHeight();
            context.RowMargin ??= GetRowMargin();
            context.LineSize ??= GetLineSize();

            var bounds = selection.GetBounds();
            var deltaX = -Math.Max(0, _scrollViewer.HorizontalOffset + context.RowFullMargin!.Value - bounds.Left);
            if (deltaX == 0)
            {
                deltaX = -Math.Min(0, _scrollViewer.HorizontalOffset + _scrollViewer.ViewportWidth - bounds.Right);
            }

            var deltaY = -Math.Max(0, _scrollViewer.VerticalOffset + context.RowFullHeight!.Value - bounds.Top);
            if (deltaY == 0)
            {
                deltaY = -Math.Min(0, _scrollViewer.VerticalOffset + _scrollViewer.ViewportHeight - bounds.Bottom);
            }

            if (deltaX == 0 && deltaY == 0)
                return false;

            _scrollViewer.ScrollToHorizontalOffset(_scrollViewer.HorizontalOffset + deltaX);
            _scrollViewer.ScrollToVerticalOffset(_scrollViewer.VerticalOffset + deltaY);
            return true;
        }

        public Rect GetCellBounds(RowCol rowCol) { ArgumentNullException.ThrowIfNull(rowCol); return new SheetSelection(this, rowCol.RowIndex, rowCol.ColumnIndex).GetBounds(); }
        public Rect GetCellBounds(int rowIndex, int columnIndex) => new SheetSelection(this, rowIndex, columnIndex).GetBounds();
        public virtual void SetColumnAutoSize(int columnIndex) => SetColumnsAutoSize(columnIndex, false, Sheet?.Rows.Values);
        private void SetColumnsAutoSize(int columnIndex, bool includeNextColumns, IEnumerable<BookDocumentRow>? rows)
        {
            if (rows == null)
                return;

            var context = CreateStyleContext();
            if (context == null || context.RowCol == null)
                throw new InvalidOperationException();

            var sizes = new Dictionary<int, double>();
            foreach (var row in rows)
            {
                context.RowCol.RowIndex = row.RowIndex;
                for (var i = columnIndex; i < _columnSettings.Count; i++)
                {
                    var colWidth = _columnSettings[i].Width;
                    var cellWidth = colWidth - (context.CellPadding.Right + CellPadding.Left);
                    if (cellWidth <= 0)
                        continue;

                    if (row.Cells.TryGetValue(i, out var cell))
                    {
                        context.RowCol.ColumnIndex = i;
                        var style = GetCellStyle(context, cell);
                        var ft = style.CreateCellFormattedText(Sheet, context, cell);
                        if (ft != null && (!sizes.TryGetValue(i, out var size) || ft.Width > size))
                        {
                            sizes[i] = ft.Width;
                        }
                    }

                    if (!includeNextColumns)
                        break;
                }
            }

            foreach (var kv in sizes)
            {
                SetColumnSize(kv.Key, kv.Value + context.CellPadding.Left + context.CellPadding.Right);
            }
        }

        protected virtual internal StyleContext CreateStyleContext()
        {
            var rowHeight = GetRowHeight();
            var cellPadding = CellPadding;
            return new()
            {
                RowCol = new RowCol(),
                MaxLineCount = 1,
                PixelsPerDip = VisualTreeHelper.GetDpi(this).PixelsPerDip,
                Typeface = new Typeface(FontFamily, FontStyle, FontWeight, FontStretch),
                FontSize = FontSize,
                Foreground = Foreground,
                TextTrimming = TextTrimming,
                RowHeight = rowHeight,
                CellHeight = rowHeight - (cellPadding.Top + cellPadding.Bottom),
                CellPadding = cellPadding,
            };
        }

        protected virtual internal void OnSelectionChanged()
        {
            _grid?.InvalidateVisual();
            RaiseEvent(new RoutedEventArgs(SelectionChangedEvent));
        }

        protected virtual void OnSheetChanged()
        {
            _columnSettings.Clear();
            var sheet = Sheet;
            if (sheet != null && sheet.LastColumnIndex.HasValue)
            {
                for (var i = 0; i <= sheet.LastColumnIndex.Value; i++)
                {
                    if (!sheet.Columns.TryGetValue(i, out var column))
                    {
                        column = new Column
                        {
                            Index = i,
                        };
                    }

                    _columnSettings.Add(new SheetControlColumn(column) { Width = GetColumnWidth() });
                }
            }

            Selection.Update();
            _grid?.InvalidateMeasure();
        }

        protected virtual BookDocumentCellStyle GetCellStyle(StyleContext context, BookDocumentCell cell)
        {
            ArgumentNullException.ThrowIfNull(context);
            if (cell is BookDocumentStyledCell styled && styled.Style != null)
                return styled.Style;

            return BookDocumentCellStyle.Empty;
        }

        protected override void OnGotKeyboardFocus(KeyboardFocusChangedEventArgs e)
        {
            Selection.Update();
            _grid?.Focus();
            _grid?.InvalidateVisual();
        }

        protected override void OnLostKeyboardFocus(KeyboardFocusChangedEventArgs e) => _grid?.InvalidateVisual();

        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();
            _scrollViewer = Template.FindName(PartScrollViewerName, this) as ScrollViewer;
            if (_scrollViewer == null)
                throw new InvalidOperationException();

            _grid = new SheetGrid(this);
            _scrollViewer.Content = _grid;
            _scrollViewer.ScrollChanged += (s, e) => _grid?.InvalidateVisual();
            RaiseEvent(new RoutedEventArgs(SelectionChangedEvent));
        }

        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            var shift = Keyboard.Modifiers.HasFlag(ModifierKeys.Shift);
            var ctl = Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
            switch (e.Key)
            {
                case Key.Escape:
                    if (IsMouseCaptured)
                    {
                        ReleaseCapture(false);
                        e.Handled = true;
                        return;
                    }
                    break;

                case Key.C:
                    if (ctl)
                    {
                        // paste unicode text + text + html
                        const string htmlClipboardHeader = "Version:0.9\r\n" + "StartHTML:{0:0000000000}\r\n" + "EndHTML:{1:0000000000}\r\n" + "StartFragment:{2:0000000000}\r\n" + "EndFragment:{3:0000000000}\r\n";
                        const string htmlClipboardStart = "<html>\r\n" + "<body>\r\n" + "<!--StartFragment-->";
                        const string htmlClipboardEnd = "<!--EndFragment-->\r\n" + "</body>\r\n" + "</html>";

                        var tl = Selection.TopLeft;
                        var br = Selection.BottomRight;
                        var sb = new StringBuilder();
                        var html = new StringBuilder(htmlClipboardHeader);
                        html.Append(htmlClipboardStart);
                        html.Append("<table>");
                        for (var i = tl.RowIndex; i <= br.RowIndex; i++)
                        {
                            var line = new StringBuilder();
                            html.Append("<tr>");
                            if (Sheet.Rows.TryGetValue(i, out var row))
                            {
                                for (var j = tl.ColumnIndex; j <= br.ColumnIndex; j++)
                                {
                                    html.Append("<td>");
                                    if (line.Length > 0)
                                    {
                                        line.Append('\t');
                                    }

                                    if (row.Cells.TryGetValue(j, out var cell))
                                    {
                                        var text = Sheet.FormatValue(cell.Value);
                                        line.Append(text);
                                        if (cell.Value != null)
                                        {
                                            html.Append(WebUtility.HtmlEncode(text));
                                        }
                                    }
                                    html.Append("</td>");
                                }
                            }
                            sb.AppendLine(line.ToString());
                            html.Append("</tr>");
                        }

                        html.Append("</table>");
                        html.Append(htmlClipboardEnd);

                        var data = new DataObject();
                        data.SetData(DataFormats.UnicodeText, sb.ToString());
                        data.SetData(DataFormats.Html, html.ToString());
                        Clipboard.SetDataObject(data);
                    }
                    break;

                case Key.A:
                    if (ctl)
                    {
                        Selection.SelectAll();
                        e.Handled = true;
                    }
                    break;

                case Key.Home:
                    Selection.MoveHorizontally(int.MinValue, shift);
                    if (ctl)
                    {
                        Selection.MoveVertically(int.MinValue, shift);
                    }
                    e.Handled = true;
                    EnsureVisible(Selection.TopLeft);
                    break;

                case Key.End:
                    Selection.MoveHorizontally(int.MaxValue, shift);
                    if (ctl)
                    {
                        Selection.MoveVertically(int.MaxValue, shift);
                    }
                    e.Handled = true;
                    EnsureVisible(Selection.BottomRight);
                    break;

                case Key.Right:
                    Selection.MoveHorizontally(ctl ? int.MaxValue : 1, shift);
                    e.Handled = true;
                    EnsureVisible(Selection.BottomRight);
                    break;

                case Key.Left:
                    Selection.MoveHorizontally(ctl ? int.MinValue : -1, shift);
                    e.Handled = true;
                    EnsureVisible(Selection.TopLeft);
                    break;

                case Key.Down:
                    Selection.MoveVertically(ctl ? int.MaxValue : 1, shift);
                    e.Handled = true;
                    EnsureVisible(Selection.BottomRight);
                    break;

                case Key.Up:
                    Selection.MoveVertically(ctl ? int.MinValue : -1, shift);
                    e.Handled = true;
                    EnsureVisible(Selection.TopLeft);
                    break;
            }
        }

        private void ReleaseCapture(bool commit)
        {
            if (_movingColumn != null && !commit)
            {
                _movingColumn.Current = _movingColumn.Start;
            }

            _movingColumn = null;
            Cursor = null;
            ReleaseMouseCapture();
        }

        protected override void OnPreviewMouseDoubleClick(MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                var sheet = Sheet;
                if (sheet != null)
                {
                    var result = HitTest(e.GetPosition(_scrollViewer));
                    if (result.MovingColumnIndex.HasValue)
                    {
                        var next = Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) && Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
                        SetColumnsAutoSize(result.MovingColumnIndex.Value, next, sheet.Rows.Values);
                    }
                }
            }
        }

        protected override void OnPreviewMouseLeftButtonUp(MouseButtonEventArgs e)
        {
            _extendingSelection = false;
            ReleaseCapture(true);
        }

        protected override void OnPreviewMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            Focus();
            var result = HitTest(e.GetPosition(_scrollViewer));
            if (result.MovingColumnIndex.HasValue)
            {
                Cursor = Cursors.SizeWE;
                CaptureMouse();
                _movingColumn = new MovingColumn(this, _columnSettings[result.MovingColumnIndex.Value], e.GetPosition(this));
                return;
            }

            if (!result.IsOverRowHeader && !result.IsOverColumnHeader && result.RowCol != null)
            {
                Selection.Select(result.RowCol);
                _extendingSelection = true;
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (_movingColumn != null)
            {
                Cursor = Cursors.SizeWE;
                _movingColumn.Current = e.GetPosition(this);
            }
            else
            {
                var result = HitTest(e.GetPosition(_scrollViewer));
                if (result.MovingColumnIndex.HasValue)
                {
                    Cursor = Cursors.SizeWE;
                }
                else
                {
                    Cursor = null;
                }

                if (_extendingSelection && !result.IsOverRowHeader && !result.IsOverColumnHeader && result.RowCol != null)
                {
                    Selection.SelectTo(result.RowCol);
                }
            }
        }

        private sealed class MovingColumn(SheetControl control, SheetControlColumn column, Point start)
        {
            private Point _current;

            public SheetControlColumn Column { get; } = column;
            public double Width { get; } = column.Width;
            public Point Start { get; } = start;
            public Point Current
            {
                get => _current;
                set
                {
                    if (_current == value)
                        return;

                    _current = value;

                    var delta = _current.X - Start.X;
                    var newWidth = Math.Max(Width + delta, _minWidth);
                    Column.Width = newWidth;
                    control._grid?.InvalidateMeasure();
                }
            }
        }

        private sealed class SheetGrid : UIElement
        {
            private bool IsSheetVisible() => _control.Sheet != null && _control.Sheet.Columns.Count > 0 && _control.Sheet.Rows.Count > 0 && _control._scrollViewer != null;

            private SheetControlHitTestResult? _lastResult;
            private Point? _lastResultPoint;
            private readonly SheetControl _control;

            public SheetGrid(SheetControl control)
            {
                _control = control;
                Focusable = true;
            }

            public SheetControlHitTestResult HitTest(Point point)
            {
                if (_lastResult != null && _lastResultPoint != null && _lastResultPoint == point)
                    return _lastResult;

                var result = new SheetControlHitTestResult();
                if (IsSheetVisible() && _control._scrollViewer != null &&
                    point.X < _control._scrollViewer.ViewportWidth &&
                    point.Y < _control._scrollViewer.ViewportHeight)
                {
                    var context = _control.CreateStyleContext();
                    if (context == null)
                        throw new InvalidOperationException();

                    context.RowHeight ??= _control.GetRowHeight();
                    context.RowMargin ??= _control.GetRowMargin();
                    context.LineSize ??= _control.GetLineSize();

                    var offsetX = _control._scrollViewer.HorizontalOffset;
                    var offsetY = _control._scrollViewer.VerticalOffset;

                    var x = point.X + offsetX;
                    var y = point.Y + offsetY;

                    var rowIndex = (int)Math.Floor((y - context.RowFullHeight!.Value) / context.RowFullHeight.Value);
                    if (rowIndex >= -1 && rowIndex <= _control.Sheet.LastRowIndex)
                    {
                        var columnIndex = GetColumnIndex(x - context.RowFullMargin!.Value, _control._scrollViewer.ExtentWidth, true);

                        // independent from scrollviewer
                        result.IsOverRowHeader = point.X >= 0 && point.X <= context.RowFullMargin.Value;
                        result.IsOverColumnHeader = point.Y >= 0 && point.Y <= context.RowFullHeight.Value;

                        if (columnIndex >= 0 || rowIndex >= 0)
                        {
                            result.RowCol = new RowCol { ColumnIndex = columnIndex ?? -1, RowIndex = rowIndex };
                        }

                        if (!result.IsOverColumnHeader && !result.IsOverRowHeader && result.RowCol != null)
                        {
                            result.Cell = _control.Sheet?.GetCell(result.RowCol);
                        }

                        if (!result.IsOverRowHeader)
                        {
                            var colSeparatorX = context.RowFullMargin.Value;
                            for (var i = 0; i < _control._columnSettings.Count; i++)
                            {
                                colSeparatorX += _control._columnSettings[i].Width;
                                if ((x + _movingColumnTolerance) >= colSeparatorX && (x - _movingColumnTolerance) <= (colSeparatorX + context.LineSize.Value))
                                {
                                    result.MovingColumnIndex = i;
                                    break;
                                }

                                colSeparatorX += context.LineSize.Value;
                            }
                        }
                    }
                }

                _lastResultPoint = point;
                _lastResult = result;
                return result;
            }

            protected override Size MeasureCore(Size availableSize)
            {
                if (!IsSheetVisible())
                    return new Size();

                var context = _control.CreateStyleContext();
                if (context == null)
                    throw new InvalidOperationException();

                context.RowHeight ??= _control.GetRowHeight();
                context.RowMargin ??= _control.GetRowMargin();
                context.LineSize ??= _control.GetLineSize();
                var columnsWidth = context.RowFullMargin!.Value; ;
                foreach (var column in _control._columnSettings)
                {
                    columnsWidth += column.Width + context.LineSize.Value;
                }

                return new Size(columnsWidth, context.RowFullHeight!.Value + context.RowFullHeight.Value * (_control.Sheet.LastRowIndex!.Value + 1));
            }

            private int? GetColumnIndex(double x, double maxWidth, bool allowReturnNull)
            {
                if (allowReturnNull && x < 0)
                    return -1;

                var lineSize = _control.GetLineSize();
                var columnsWidth = lineSize;
                foreach (var column in _control._columnSettings)
                {
                    columnsWidth += column.Width + lineSize;
                    if (columnsWidth > maxWidth)
                        return column.Column.Index;

                    if (x < columnsWidth)
                        return column.Column.Index;
                }
                return allowReturnNull ? null : _control.Sheet.LastColumnIndex;
            }

            protected override void OnRender(DrawingContext drawingContext)
            {
                base.OnRender(drawingContext);
                _lastResult = null;
                _lastResultPoint = null;
                if (!IsSheetVisible())
                    return;

                var viewWidth = _control._scrollViewer!.ViewportWidth;
                if (viewWidth == 0)
                    return;

                var viewHeight = _control._scrollViewer.ViewportHeight;
                if (viewHeight == 0)
                    return;

                var context = _control.CreateStyleContext();
                if (context == null || context.RowCol == null)
                    throw new InvalidOperationException();

                context.RowHeight ??= _control.GetRowHeight();
                context.RowMargin ??= _control.GetRowMargin();
                context.LineSize ??= _control.GetLineSize();
                context.LinePen ??= new Pen(_control.LineBrush, context.LineSize.Value);

                var offsetX = _control._scrollViewer.HorizontalOffset;
                var offsetY = _control._scrollViewer.VerticalOffset;

                var firstDrawnColumnIndex = GetColumnIndex(offsetX - context.RowFullMargin!.Value, _control._scrollViewer.ExtentWidth, false);
                var lastDrawnColumnIndex = GetColumnIndex(offsetX + viewWidth - context.RowFullMargin!.Value, _control._scrollViewer.ExtentWidth, false);
                if (!firstDrawnColumnIndex.HasValue || !lastDrawnColumnIndex.HasValue)
                    return;

                var firstDrawnRowIndex = Math.Max((int)((offsetY - context.RowFullHeight!.Value) / context.RowFullHeight!.Value), 0);
                var lastDrawnRowIndex = Math.Max(Math.Min((int)((offsetY - context.RowFullHeight.Value + viewHeight) / context.RowFullHeight.Value), _control.Sheet.LastRowIndex!.Value), firstDrawnRowIndex);

                // compute first col X
                var startCurrentColX = context.RowMargin.Value + context.LineSize.Value / 2;
                for (var k = 0; k < firstDrawnColumnIndex.Value; k++)
                {
                    startCurrentColX += _control._columnSettings[k].Width + context.LineSize.Value;
                }

                // draw cells
                double currentRowY;
                var cellPadding = _control.CellPadding;
                var cellHeight = context.RowHeight.Value - (cellPadding.Top + cellPadding.Bottom);
                if (cellHeight > 0)
                {
                    var cellsRect = new Rect(offsetX + context.RowMargin.Value, offsetY + context.RowHeight.Value, viewWidth - context.RowMargin.Value, viewHeight - context.RowHeight.Value);
                    drawingContext.PushClip(new RectangleGeometry(cellsRect));

                    currentRowY = context.RowHeight.Value + context.LineSize.Value / 2 + context.RowFullHeight.Value * firstDrawnRowIndex;
                    for (var i = firstDrawnRowIndex; i <= lastDrawnRowIndex; i++)
                    {
                        if (_control.Sheet.Rows.TryGetValue(i, out var row))
                        {
                            context.RowCol.RowIndex = i;
                            var ccx = startCurrentColX;
                            for (var j = firstDrawnColumnIndex.Value; j <= lastDrawnColumnIndex; j++)
                            {
                                var colWidth = _control._columnSettings[j].Width;
                                var cellWidth = colWidth - (cellPadding.Right + cellPadding.Left);
                                if (cellWidth <= 0)
                                    continue;

                                if (row.Cells.TryGetValue(j, out var cell))
                                {
                                    context.RowCol.RowIndex = j;
                                    context.ColumnWidth = colWidth;
                                    context.CellWidth = cellWidth;
                                    var style = _control.GetCellStyle(context, cell);

                                    if (style.Background != null)
                                    {
                                        drawingContext.DrawRectangle(style.Background, null, new Rect(ccx, currentRowY, colWidth, cellHeight));
                                    }

                                    var formattedCell = style.CreateCellFormattedText(_control.Sheet, context, cell);
                                    if (formattedCell != null)
                                    {
                                        var textOffsetY = (context.RowHeight.Value - formattedCell.Height) / 2; // center vertically
                                        drawingContext.DrawText(formattedCell, new Point(ccx + cellPadding.Left + context.LineSize.Value / 2, currentRowY + context.LineSize.Value / 2 + textOffsetY));
                                    }
                                }
                                ccx += colWidth + context.LineSize.Value;
                            }
                        }
                        currentRowY += context.RowFullHeight.Value;
                    }
                    drawingContext.Pop();
                }

                var rowsHeight = (lastDrawnRowIndex + 1) * context.RowFullHeight.Value + context.RowFullHeight.Value;
                var columnsWidth = context.RowFullMargin!.Value;
                foreach (var column in _control._columnSettings)
                {
                    columnsWidth += context.LineSize.Value + column.Width;
                    if (columnsWidth >= viewWidth)
                    {
                        columnsWidth = viewWidth;
                        break;
                    }
                }

                // draw rows
                var rowsRect = new Rect(offsetX, offsetY + context.RowHeight.Value, viewWidth, Math.Min(viewHeight, rowsHeight) - context.RowHeight.Value);
                drawingContext.PushClip(new RectangleGeometry(rowsRect));
                rowsRect.Width = context.RowMargin.Value;
                drawingContext.DrawRectangle(_control.HeaderBrush, null, rowsRect);

                currentRowY = context.RowHeight.Value + context.LineSize.Value / 2 + context.RowFullHeight.Value * firstDrawnRowIndex;
                for (var i = firstDrawnRowIndex; i <= lastDrawnRowIndex; i++)
                {
                    // draw row name
                    var name = (i + 1).ToString();
                    var formattedRow = new FormattedText(name, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, context.Typeface, _control.FontSize, _control.Foreground, context.PixelsPerDip)
                    {
                        MaxTextWidth = context.RowMargin.Value,
                        MaxLineCount = 1
                    };

                    var textOffsetX = (context.RowMargin.Value - formattedRow.Width) / 2; // center horizontally
                    var textOffsetY = (context.RowHeight.Value - formattedRow.Height) / 2; // center vertically
                    drawingContext.DrawText(formattedRow, new Point(offsetX + textOffsetX, currentRowY + context.LineSize.Value / 2 + textOffsetY));

                    drawingContext.DrawLine(context.LinePen, new Point(offsetX, currentRowY), new Point(offsetX + columnsWidth, currentRowY));
                    currentRowY += context.RowFullHeight.Value;
                }

                // last row
                drawingContext.DrawLine(context.LinePen, new Point(offsetX, currentRowY), new Point(offsetX + columnsWidth, currentRowY));
                drawingContext.Pop();

                // draw columns
                var columnsRect = new Rect(offsetX + context.RowMargin.Value, offsetY, columnsWidth - context.RowMargin.Value, viewHeight);
                drawingContext.PushClip(new RectangleGeometry(columnsRect));
                columnsRect.Height = context.RowHeight.Value;
                drawingContext.DrawRectangle(_control.HeaderBrush, null, columnsRect);

                var currentColX = startCurrentColX;
                for (var i = firstDrawnColumnIndex.Value; i <= lastDrawnColumnIndex; i++)
                {
                    // draw col name
                    var colWidth = _control._columnSettings[i].Width;

                    _control.Sheet.Columns.TryGetValue(i, out var col);
                    var name = col?.Name ?? Row.GetExcelColumnName(i);
                    var formattedCol = new FormattedText(name, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, context.Typeface, _control.FontSize, _control.Foreground, context.PixelsPerDip)
                    {
                        MaxTextWidth = colWidth,
                        MaxLineCount = 1
                    };

                    var textOffsetX = (colWidth - formattedCol.Width) / 2; // center horizontally
                    var textOffsetY = (context.RowHeight.Value - formattedCol.Height) / 2; // center vertically
                    drawingContext.DrawText(formattedCol, new Point(currentColX + context.LineSize.Value / 2 + textOffsetX, offsetY + textOffsetY));

                    drawingContext.DrawLine(context.LinePen, new Point(currentColX, offsetY), new Point(currentColX, offsetY + Math.Min(rowsHeight, viewHeight)));
                    currentColX += colWidth + context.LineSize.Value;
                }

                // last col
                drawingContext.DrawLine(context.LinePen, new Point(currentColX, offsetY), new Point(currentColX, offsetY + Math.Min(rowsHeight, viewHeight)));
                drawingContext.Pop();

                // focus & selection
                var selectionBrush = _control.SelectionBrush;
                if (selectionBrush != null)
                {
                    var focusMargin = context.LineSize.Value + 2;
                    var cellsRect = new Rect(
                        offsetX + context.RowMargin.Value - focusMargin,
                        offsetY + context.RowHeight.Value - focusMargin,
                        viewWidth - context.RowMargin.Value + focusMargin,
                        viewHeight - context.RowHeight.Value + focusMargin);
                    drawingContext.PushClip(new RectangleGeometry(cellsRect));
                    var rc = _control.Selection.GetBounds();
                    if (IsKeyboardFocused)
                    {
                        var focusRc = rc;
                        focusRc.Inflate(focusMargin, focusMargin);
                        var focusPen = new Pen(selectionBrush, context.LineSize.Value + 1) { DashStyle = new DashStyle([0.0, 3.0], 0) };
                        drawingContext.DrawRectangle(null, focusPen, focusRc);
                    }

                    var pen = new Pen(selectionBrush, context.LineSize.Value + 1);
                    drawingContext.DrawRectangle(null, pen, rc);
                    drawingContext.Pop();
                }
            }
        }
    }
}
