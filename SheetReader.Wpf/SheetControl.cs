using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
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

        private const double _minWidth = 4;
        private const double _movingColumnTolerance = 4;

        private double GetRowHeight() => Math.Max(RowHeight, _minWidth);
        private double GetRowMargin() => Math.Max(RowMargin, _minWidth);
        private double GetLineSize() => Math.Max(LineSize, 0);
        private double GetColumnWidth() => Math.Max(ColumnWidth, _minWidth);

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
            Focusable = false;
            Selection = new SheetSelection(this);
        }

        public SheetSelection Selection { get; }

        public virtual SheetControlHitTestResult HitTest(Point point) => _grid?.HitTest(point) ?? new();
        public virtual void SetColumnSize(int columnIndex, double width)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(width);
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columnIndex, _columnSettings.Count);
            _columnSettings[columnIndex].Width = width;
            _grid?.InvalidateMeasure();
        }

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
                        var ft = style.CreateCellFormattedText(context, cell);
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

        protected virtual StyleContext CreateStyleContext()
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

        protected virtual internal void OnSelectionChanged() => _grid?.InvalidateVisual();
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
            Log("OnGotKeyboardFocus");
            base.OnGotKeyboardFocus(e);
        }

        protected override void OnLostKeyboardFocus(KeyboardFocusChangedEventArgs e)
        {
            Log("OnLostKeyboardFocus");
            base.OnLostKeyboardFocus(e);
        }

        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();
            _scrollViewer = Template.FindName(PartScrollViewerName, this) as ScrollViewer;
            if (_scrollViewer == null)
                throw new InvalidOperationException();

            _grid = new SheetGrid(this);
            _scrollViewer.Content = _grid;
            _scrollViewer.ScrollChanged += (s, e) => _grid?.InvalidateVisual();
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
                        return;
                    }
                    break;

                case Key.A:
                    if (ctl)
                    {
                        Selection.SelectAll();
                    }
                    break;

                case Key.Home:
                    Selection.MoveHorizontally(int.MinValue, shift);
                    if (ctl)
                    {
                        Selection.MoveVertically(int.MinValue, shift);
                    }
                    break;

                case Key.End:
                    Selection.MoveHorizontally(int.MaxValue, shift);
                    if (ctl)
                    {
                        Selection.MoveVertically(int.MaxValue, shift);
                    }
                    break;

                case Key.Right:
                    Selection.MoveHorizontally(ctl ? int.MaxValue : 1, shift);
                    break;

                case Key.Left:
                    Selection.MoveHorizontally(ctl ? int.MinValue : -1, shift);
                    break;

                case Key.Down:
                    Selection.MoveVertically(ctl ? int.MaxValue : 1, shift);
                    break;

                case Key.Up:
                    Selection.MoveVertically(ctl ? int.MinValue : -1, shift);
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

        private sealed class SheetGrid(SheetControl control) : UIElement
        {
            private bool IsSheetVisible() => _control.Sheet != null && _control.Sheet.Columns.Count > 0 && _control.Sheet.Rows.Count > 0 && _control._scrollViewer != null;

            private SheetControlHitTestResult? _lastResult;
            private Point? _lastResultPoint;
            private readonly SheetControl _control = control;

            public SheetControlHitTestResult HitTest(Point point)
            {
                if (_lastResult != null && _lastResultPoint != null && _lastResultPoint == point)
                    return _lastResult;

                var result = new SheetControlHitTestResult();
                if (IsSheetVisible())
                {
                    var lineSize = _control.GetLineSize();
                    var rowHeight = _control.GetRowHeight();
                    var rowFullHeight = rowHeight + lineSize;
                    var rowsHeaderWidth = _control.GetRowMargin();
                    var rowsHeaderFullWidth = rowsHeaderWidth + lineSize;
                    var columnsHeaderHeight = rowHeight;
                    var columnsHeaderFullHeight = columnsHeaderHeight + lineSize;

                    var offsetX = _control._scrollViewer!.HorizontalOffset;
                    var offsetY = _control._scrollViewer.VerticalOffset;

                    var x = point.X + offsetX;
                    var y = point.Y + offsetY;

                    var rowIndex = (int)Math.Floor((y - columnsHeaderFullHeight) / rowFullHeight);
                    if (rowIndex >= -1 && rowIndex <= _control.Sheet.LastRowIndex)
                    {
                        var columnIndex = GetColumnIndex(x - rowsHeaderFullWidth, _control._scrollViewer!.ExtentWidth, true);

                        // independent from scrollviewer
                        result.IsOverRowHeader = point.X >= 0 && point.X <= rowsHeaderFullWidth;
                        result.IsOverColumnHeader = point.Y >= 0 && point.Y <= columnsHeaderFullHeight;

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
                            var colSeparatorX = rowsHeaderFullWidth;
                            for (var i = 0; i < _control._columnSettings.Count; i++)
                            {
                                colSeparatorX += _control._columnSettings[i].Width;
                                if ((x + _movingColumnTolerance) >= colSeparatorX && (x - _movingColumnTolerance) <= (colSeparatorX + lineSize))
                                {
                                    result.MovingColumnIndex = i;
                                    break;
                                }

                                colSeparatorX += lineSize;
                            }
                        }
                    }

                    //Log("hit: " + result.RowCol + " (" + result.RowCol?.ExcelReference + ") row:" + result.IsOverRowHeader + " column:" + result.IsOverColumnHeader + " moving:" + result.MovingColumnIndex);
                }

                _lastResultPoint = point;
                _lastResult = result;
                return result;
            }

            protected override Size MeasureCore(Size availableSize)
            {
                if (!IsSheetVisible())
                    return new Size();

                var lineSize = _control.GetLineSize();
                var rowHeight = _control.GetRowHeight();
                var rowFullHeight = rowHeight + lineSize;
                var rowsHeaderWidth = _control.GetRowMargin();
                var columnsHeaderHeight = rowHeight;
                var columnsHeaderFullHeight = columnsHeaderHeight + lineSize;

                var columnsWidth = lineSize + rowsHeaderWidth;
                foreach (var column in _control._columnSettings)
                {
                    columnsWidth += column.Width + lineSize;
                }

                return new Size(columnsWidth, columnsHeaderFullHeight + rowFullHeight * (_control.Sheet.LastRowIndex!.Value + 1));
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
                //Log("offset:" + offsetX + " x " + offsetY + " view:" + viewWidth + " x " + viewHeight);

                var firstDrawnColumnIndex = GetColumnIndex(offsetX - context.RowFullMargin!.Value, _control._scrollViewer.ExtentWidth, false);
                var lastDrawnColumnIndex = GetColumnIndex(offsetX + viewWidth - context.RowFullMargin!.Value, _control._scrollViewer.ExtentWidth, false);
                if (!firstDrawnColumnIndex.HasValue || !lastDrawnColumnIndex.HasValue)
                    return;

                var firstDrawnRowIndex = Math.Max((int)((offsetY - context.RowFullHeight!.Value) / context.RowFullHeight!.Value), 0);
                var lastDrawnRowIndex = Math.Max(Math.Min((int)((offsetY - context.RowFullHeight.Value + viewHeight) / context.RowFullHeight.Value), _control.Sheet.LastRowIndex!.Value), firstDrawnRowIndex);

                //Log("col:" + firstDrawnColumnIndex + " => " + lastDrawnColumnIndex + " row:" + firstDrawnRowIndex + " => " + lastDrawnRowIndex);

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
                        if (!_control.Sheet.Rows.TryGetValue(i, out var row))
                            continue;

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
                                var formattedCell = style.CreateCellFormattedText(context, cell);
                                if (formattedCell != null)
                                {
                                    var textOffsetY = (context.RowHeight.Value - formattedCell.Height) / 2; // center vertically
                                    drawingContext.DrawText(formattedCell, new Point(ccx + cellPadding.Left + context.LineSize.Value / 2, currentRowY + context.LineSize.Value / 2 + textOffsetY));
                                }
                            }
                            ccx += colWidth + context.LineSize.Value;
                            //Log("trace " + i + " x " + j + " " + text);
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
                    var name = Row.GetExcelColumnName(i);
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

                var selectionBrush = _control.SelectionBrush;
                if (selectionBrush != null)
                {
                    var pen = new Pen(selectionBrush, context.LineSize.Value * 3);

                    var x = context.RowFullMargin.Value;
                    for (var i = 0; i < _control.Selection.ColumnIndex; i++)
                    {
                        x += _control._columnSettings[i].Width + context.LineSize.Value;
                    }

                    var w = _control._columnSettings[_control.Selection.ColumnIndex].Width + context.LineSize.Value;
                    if (_control.Selection.ColumnExtension < 0)
                    {
                        for (var i = 1; i <= -_control.Selection.ColumnExtension; i++)
                        {
                            var width = _control._columnSettings[_control.Selection.ColumnIndex - i].Width + context.LineSize.Value;
                            x -= width;
                            w += width;
                        }
                    }
                    else
                    {
                        for (var i = 1; i <= _control.Selection.ColumnExtension; i++)
                        {
                            w += _control._columnSettings[_control.Selection.ColumnIndex + i].Width + context.LineSize.Value;
                        }
                    }

                    var y = context.RowFullHeight.Value + context.RowFullHeight.Value * _control.Selection.RowIndex;
                    var h = context.RowFullHeight.Value;
                    if (_control.Selection.RowExtension < 0)
                    {
                        var height = -_control.Selection.RowExtension * context.RowFullHeight.Value;
                        h += height;
                        y -= height;
                    }
                    else
                    {
                        h += _control.Selection.RowExtension * context.RowFullHeight.Value;
                    }
                    var rc = new Rect(x, y, w, h);
                    drawingContext.DrawRectangle(null, pen, rc);
                }
            }
        }
    }
}
