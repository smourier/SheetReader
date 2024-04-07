using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
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
            new UIPropertyMetadata(new Thickness(5)));

        public BookDocumentSheet Sheet { get => (BookDocumentSheet)GetValue(SheetProperty); set => SetValue(SheetProperty, value); }
        public double ColumnWidth { get => (double)GetValue(ColumnWidthProperty); set => SetValue(ColumnWidthProperty, value); }
        public double RowMargin { get => (double)GetValue(RowMarginProperty); set => SetValue(RowMarginProperty, value); }
        public double RowHeight { get => (double)GetValue(RowHeightProperty); set => SetValue(RowHeightProperty, value); }
        public double LineSize { get => (double)GetValue(LineSizeProperty); set => SetValue(LineSizeProperty, value); }
        public Brush LineBrush { get => (Brush)GetValue(LineBrushProperty); set => SetValue(LineBrushProperty, value); }
        public Brush HeaderBrush { get => (Brush)GetValue(HeaderBrushProperty); set => SetValue(HeaderBrushProperty, value); }
        public TextAlignment TextAlignment { get => (TextAlignment)GetValue(TextAlignmentProperty); set => SetValue(TextAlignmentProperty, value); }
        public TextTrimming TextTrimming { get => (TextTrimming)GetValue(TextTrimmingProperty); set => SetValue(TextTrimmingProperty, value); }
        public Thickness CellPadding { get => (Thickness)GetValue(CellPaddingProperty); set => SetValue(CellPaddingProperty, value); }

        private double GetRowHeight() => Math.Max(RowHeight, 10);
        private double GetRowMargin() => Math.Max(RowMargin, 10);
        private double GetLineSize() => Math.Max(LineSize, 0);

        static SheetControl()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(SheetControl), new FrameworkPropertyMetadata(typeof(SheetControl)));
        }

        private static void Log(object? message = null, [CallerMemberName] string? methodName = null) => EventProvider.Default.WriteMessageEvent(methodName + ":" + message);

        private ScrollViewer? _scrollViewer;
        private SheetGrid? _grid;
        private MovingColumn? _movingColumn;
        private readonly List<SheetControlColumn> _columnSettings = [];

        public virtual SheetControlHitTestResult HitTest(Point point) => _grid?.HitTest(point) ?? new();
        public virtual void SetColumnSize(int columnIndex, double width)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(width);
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columnIndex, _columnSettings.Count);
            _columnSettings[columnIndex].Width = width;
            _grid?.InvalidateMeasure();
            Log("colsize " + columnIndex + ": " + width);
        }

        public virtual void SetColumnAutoSize(int columnIndex)
        {
            var typeface = new Typeface(FontFamily, FontStyle, FontWeight, FontStretch);
            SetColumnAutoSize(columnIndex, Sheet?.Rows.Values, typeface, FontSize, Foreground, TextTrimming, VisualTreeHelper.GetDpi(this).PixelsPerDip);
        }

        private void SetColumnAutoSize(int columnIndex, IEnumerable<BookDocumentRow>? rows, Typeface typeface, double fontSize, Brush foreground, TextTrimming textTrimming, double pixelsPerDip)
        {
            if (rows == null)
                return;

            double? autoSize = null;
            foreach (var row in rows)
            {
                if (row.Cells.TryGetValue(columnIndex, out var cell))
                {
                    var text = string.Format("{0}", cell.Value);
                    if (!string.IsNullOrEmpty(text))
                    {
                        var ft = new FormattedText(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, typeface, fontSize, foreground, pixelsPerDip)
                        {
                            Trimming = textTrimming,
                            MaxLineCount = 1
                        };

                        if (!autoSize.HasValue || ft.WidthIncludingTrailingWhitespace > autoSize)
                        {
                            autoSize = ft.WidthIncludingTrailingWhitespace;
                        }
                    }
                }
            }

            if (autoSize.HasValue)
            {
                Dispatcher.BeginInvoke(() =>
                {
                    SetColumnSize(columnIndex, autoSize.Value);
                });
            }
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

                    _columnSettings.Add(new SheetControlColumn(column) { Width = ColumnWidth });
                }
            }

            _grid?.InvalidateMeasure();
        }

        protected virtual BookDocumentCellStyle GetCellStyle(GetStyleContext context, BookDocumentCell cell)
        {
            ArgumentNullException.ThrowIfNull(context);
            if (cell is BookDocumentStyledCell styled && styled.Style != null)
                return styled.Style;

            return new BookDocumentCellStyle();
        }

        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();
            _scrollViewer = Template.FindName(PartScrollViewerName, this) as ScrollViewer;
            if (_scrollViewer == null)
                throw new InvalidOperationException();

            _grid = new SheetGrid(this);
            _scrollViewer.Content = _grid;
            _scrollViewer.ScrollChanged += OnScrollChanged;
        }

        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                ReleaseCapture(false);
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
                    var typeface = new Typeface(FontFamily, FontStyle, FontWeight, FontStretch);
                    var fg = Foreground;
                    var fs = FontSize;
                    var tt = TextTrimming;
                    var ppd = VisualTreeHelper.GetDpi(this).PixelsPerDip;
                    var result = HitTest(e.GetPosition(_scrollViewer));
                    if (result.MovingColumnIndex.HasValue)
                    {
                        for (var i = result.MovingColumnIndex.Value; i <= sheet.LastColumnIndex; i++)
                        {
                            var j = i;
                            Task.Run(() =>
                            {
                                SetColumnAutoSize(j, sheet.Rows.Values, typeface, fs, fg, tt, ppd);
                            });
                        }
                    }
                }
            }
        }

        protected override void OnPreviewMouseLeftButtonUp(MouseButtonEventArgs e) => ReleaseCapture(true);
        protected override void OnPreviewMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            if (_scrollViewer == null)
                return;

            var result = HitTest(e.GetPosition(_scrollViewer));
            if (result.MovingColumnIndex.HasValue)
            {
                Cursor = Cursors.SizeWE;
                CaptureMouse();
                _movingColumn = new MovingColumn(this, _columnSettings[result.MovingColumnIndex.Value], e.GetPosition(this));
                Focus();
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (_scrollViewer == null)
                return;

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
            }
        }

        private void OnScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            _grid?.InvalidateVisual();
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

                    const int minSize = 4;
                    var delta = _current.X - Start.X;
                    var newWidth = Math.Max(Width + delta, minSize);
                    Column.Width = newWidth;
                    control._grid?.InvalidateMeasure();
                }
            }
        }

        private sealed class SheetGrid(SheetControl control) : UIElement
        {
            private bool IsSheetVisible() => control.Sheet != null && control.Sheet.Columns.Count > 0 && control.Sheet.Rows.Count > 0 && control._scrollViewer != null;

            private SheetControlHitTestResult? _lastResult;
            private Point? _lastResultPoint;

            public SheetControlHitTestResult HitTest(Point point)
            {
                if (_lastResult != null && _lastResultPoint != null && _lastResultPoint == point)
                    return _lastResult;

                var result = new SheetControlHitTestResult();
                if (IsSheetVisible())
                {
                    var lineSize = control.GetLineSize();
                    var rowHeight = control.GetRowHeight();
                    var rowFullHeight = rowHeight + lineSize;
                    var rowsHeaderWidth = control.GetRowMargin();
                    var rowsHeaderFullWidth = rowsHeaderWidth + lineSize;
                    var columnsHeaderHeight = rowHeight;
                    var columnsHeaderFullHeight = columnsHeaderHeight + lineSize;

                    var offsetX = control._scrollViewer!.HorizontalOffset;
                    var offsetY = control._scrollViewer.VerticalOffset;

                    var x = point.X + offsetX;
                    var y = point.Y + offsetY;

                    var rowIndex = (int)Math.Floor((y - columnsHeaderFullHeight) / rowFullHeight);
                    if (rowIndex >= -1 && rowIndex <= control.Sheet.LastRowIndex)
                    {
                        var columnIndex = GetColumnIndex(x - rowsHeaderFullWidth, control._scrollViewer!.ExtentWidth, true);

                        // independent from scrollviewer
                        result.IsOverRowHeader = point.X >= 0 && point.X <= rowsHeaderFullWidth;
                        result.IsOverColumnHeader = point.Y >= 0 && point.Y <= columnsHeaderFullHeight;

                        if (columnIndex >= 0 || rowIndex >= 0)
                        {
                            result.RowCol = new RowCol { ColumnIndex = columnIndex ?? -1, RowIndex = rowIndex };
                        }

                        if (!result.IsOverColumnHeader && !result.IsOverRowHeader && result.RowCol != null)
                        {
                            result.Cell = control.Sheet?.GetCell(result.RowCol);
                        }

                        if (!result.IsOverRowHeader)
                        {
                            const int tolerance = 4;
                            var colSeparatorX = rowsHeaderFullWidth;
                            for (var i = 0; i < control._columnSettings.Count; i++)
                            {
                                colSeparatorX += control._columnSettings[i].Width;
                                if ((x + tolerance) >= colSeparatorX && (x - tolerance) <= (colSeparatorX + lineSize))
                                {
                                    result.MovingColumnIndex = i;
                                    break;
                                }

                                colSeparatorX += lineSize;
                            }
                        }
                    }

                    Log("hit: " + result.RowCol + " (" + result.RowCol?.ExcelReference + ") row:" + result.IsOverRowHeader + " column:" + result.IsOverColumnHeader + " moving:" + result.MovingColumnIndex);
                }

                _lastResultPoint = point;
                _lastResult = result;
                return result;
            }

            protected override Size MeasureCore(Size availableSize)
            {
                if (!IsSheetVisible())
                    return new Size();

                var lineSize = control.GetLineSize();
                var rowHeight = control.GetRowHeight();
                var rowFullHeight = rowHeight + lineSize;
                var rowsHeaderWidth = control.GetRowMargin();
                var columnsHeaderHeight = rowHeight;
                var columnsHeaderFullHeight = columnsHeaderHeight + lineSize;

                var columnsWidth = lineSize + rowsHeaderWidth;
                foreach (var column in control._columnSettings)
                {
                    columnsWidth += column.Width + lineSize;
                }

                return new Size(columnsWidth, columnsHeaderFullHeight + rowFullHeight * (control.Sheet.LastRowIndex!.Value + 1));
            }

            private int? GetColumnIndex(double x, double maxWidth, bool allowReturnNull)
            {
                if (allowReturnNull && x < 0)
                    return -1;

                var lineSize = control.GetLineSize();
                var columnsWidth = lineSize;
                foreach (var column in control._columnSettings)
                {
                    columnsWidth += column.Width + lineSize;
                    if (columnsWidth > maxWidth)
                        return column.Column.Index;

                    if (x < columnsWidth)
                        return column.Column.Index;
                }
                return allowReturnNull ? null : control.Sheet.LastColumnIndex;
            }

            protected override void OnRender(DrawingContext drawingContext)
            {
                _lastResult = null;
                _lastResultPoint = null;
                if (!IsSheetVisible())
                    return;

                var viewWidth = control._scrollViewer!.ViewportWidth;
                if (viewWidth == 0)
                    return;

                var viewHeight = control._scrollViewer.ViewportHeight;
                if (viewHeight == 0)
                    return;

                var lineSize = control.GetLineSize();
                var pen = new Pen(control.LineBrush, lineSize);
                var typeface = new Typeface(control.FontFamily, control.FontStyle, control.FontWeight, control.FontStretch);
                var dpi = VisualTreeHelper.GetDpi(this);
                var rowHeight = control.GetRowHeight();
                var rowFullHeight = rowHeight + lineSize;
                var rowsHeaderWidth = control.GetRowMargin();
                var rowsHeaderFullWidth = rowsHeaderWidth + lineSize;
                var columnsHeaderHeight = rowHeight;
                var columnsHeaderFullHeight = columnsHeaderHeight + lineSize;

                var offsetX = control._scrollViewer.HorizontalOffset;
                var offsetY = control._scrollViewer.VerticalOffset;
                //Log("offset:" + offsetX + " x " + offsetY + " view:" + viewWidth + " x " + viewHeight);

                var firstDrawnColumnIndex = GetColumnIndex(offsetX - rowsHeaderFullWidth, control._scrollViewer.ExtentWidth, false);
                var lastDrawnColumnIndex = GetColumnIndex(offsetX + viewWidth - rowsHeaderFullWidth, control._scrollViewer.ExtentWidth, false);
                if (!firstDrawnColumnIndex.HasValue || !lastDrawnColumnIndex.HasValue)
                    return;

                var firstDrawnRowIndex = Math.Max((int)((offsetY - columnsHeaderFullHeight) / rowFullHeight), 0);
                var lastDrawnRowIndex = Math.Max(Math.Min((int)((offsetY - columnsHeaderFullHeight + viewHeight) / rowFullHeight), control.Sheet.LastRowIndex!.Value), firstDrawnRowIndex);

                //Log("col:" + firstDrawnColumnIndex + " => " + lastDrawnColumnIndex + " row:" + firstDrawnRowIndex + " => " + lastDrawnRowIndex);

                // compute first col X
                var startCurrentColX = rowsHeaderWidth + lineSize / 2;
                for (var k = 0; k < firstDrawnColumnIndex.Value; k++)
                {
                    startCurrentColX += control._columnSettings[k].Width + lineSize;
                }

                // draw cells
                var cellsRect = new Rect(offsetX + rowsHeaderWidth, offsetY + columnsHeaderHeight, viewWidth - rowsHeaderWidth, viewHeight - columnsHeaderHeight);
                drawingContext.PushClip(new RectangleGeometry(cellsRect));

                var context = new GetStyleContext
                {
                    RowCol = new RowCol(),
                    MaxLineCount = 1,
                    PixelsPerDip = dpi.PixelsPerDip,
                    Typeface = typeface,
                    FontSize = control.FontSize,
                    Foreground = control.Foreground,
                    TextTrimming = control.TextTrimming,
                };
                var currentRowY = columnsHeaderHeight + lineSize / 2 + rowFullHeight * firstDrawnRowIndex;
                for (var i = firstDrawnRowIndex; i <= lastDrawnRowIndex; i++)
                {
                    if (!control.Sheet.Rows.TryGetValue(i, out var row))
                        continue;

                    context.RowCol.RowIndex = i;
                    var ccx = startCurrentColX;
                    for (var j = firstDrawnColumnIndex.Value; j <= lastDrawnColumnIndex; j++)
                    {
                        var colWidth = control._columnSettings[j].Width;
                        context.ColumnWidth = colWidth;
                        if (row.Cells.TryGetValue(j, out var cell))
                        {
                            context.RowCol.RowIndex = j;
                            var style = control.GetCellStyle(context, cell);
                            var formattedCell = style.CreateCellFormattedText(context, cell);
                            if (formattedCell != null)
                            {
                                var textOffsetY = (rowHeight - formattedCell.Height) / 2; // center vertically
                                drawingContext.DrawText(formattedCell, new Point(ccx + lineSize / 2, currentRowY + lineSize / 2 + textOffsetY));
                            }
                        }
                        ccx += colWidth + lineSize;
                        //Log("trace " + i + " x " + j + " " + text);
                    }
                    currentRowY += rowFullHeight;
                }
                drawingContext.Pop();

                var rowsHeight = (lastDrawnRowIndex + 1) * rowFullHeight + columnsHeaderFullHeight;
                var columnsWidth = rowsHeaderFullWidth;
                foreach (var column in control._columnSettings)
                {
                    columnsWidth += lineSize + column.Width;
                    if (columnsWidth >= viewWidth)
                    {
                        columnsWidth = viewWidth;
                        break;
                    }
                }

                // draw rows
                var rowsRect = new Rect(offsetX, offsetY + columnsHeaderHeight, viewWidth, Math.Min(viewHeight, rowsHeight) - columnsHeaderHeight);
                drawingContext.PushClip(new RectangleGeometry(rowsRect));
                rowsRect.Width = rowsHeaderWidth;
                drawingContext.DrawRectangle(control.HeaderBrush, null, rowsRect);

                currentRowY = columnsHeaderHeight + lineSize / 2 + rowFullHeight * firstDrawnRowIndex;
                for (var i = firstDrawnRowIndex; i <= lastDrawnRowIndex; i++)
                {
                    // draw row name
                    var name = (i + 1).ToString();
                    var formattedRow = new FormattedText(name, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, typeface, control.FontSize, control.Foreground, dpi.PixelsPerDip)
                    {
                        MaxTextWidth = rowsHeaderWidth,
                        MaxLineCount = 1
                    };

                    var textOffsetX = (rowsHeaderWidth - formattedRow.Width) / 2; // center horizontally
                    var textOffsetY = (rowHeight - formattedRow.Height) / 2; // center vertically
                    drawingContext.DrawText(formattedRow, new Point(offsetX + textOffsetX, currentRowY + lineSize / 2 + textOffsetY));

                    drawingContext.DrawLine(pen, new Point(offsetX, currentRowY), new Point(offsetX + columnsWidth, currentRowY));
                    currentRowY += rowFullHeight;
                }

                // last row
                drawingContext.DrawLine(pen, new Point(offsetX, currentRowY), new Point(offsetX + columnsWidth, currentRowY));
                drawingContext.Pop();

                // draw columns
                var columnsRect = new Rect(offsetX + rowsHeaderWidth, offsetY, columnsWidth - rowsHeaderWidth, viewHeight);
                drawingContext.PushClip(new RectangleGeometry(columnsRect));
                columnsRect.Height = columnsHeaderHeight;
                drawingContext.DrawRectangle(control.HeaderBrush, null, columnsRect);

                var currentColX = startCurrentColX;
                for (var i = firstDrawnColumnIndex.Value; i <= lastDrawnColumnIndex; i++)
                {
                    // draw col name
                    var colWidth = control._columnSettings[i].Width;
                    var name = Row.GetExcelColumnName(i);
                    var formattedCol = new FormattedText(name, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, typeface, control.FontSize, control.Foreground, dpi.PixelsPerDip)
                    {
                        MaxTextWidth = colWidth,
                        MaxLineCount = 1
                    };

                    var textOffsetX = (colWidth - formattedCol.Width) / 2; // center horizontally
                    var textOffsetY = (columnsHeaderHeight - formattedCol.Height) / 2; // center vertically
                    drawingContext.DrawText(formattedCol, new Point(currentColX + lineSize / 2 + textOffsetX, offsetY + textOffsetY));

                    drawingContext.DrawLine(pen, new Point(currentColX, offsetY), new Point(currentColX, offsetY + Math.Min(rowsHeight, viewHeight)));
                    currentColX += colWidth + lineSize;
                }

                // last col
                drawingContext.DrawLine(pen, new Point(currentColX, offsetY), new Point(currentColX, offsetY + Math.Min(rowsHeight, viewHeight)));
                drawingContext.Pop();
            }
        }
    }
}
