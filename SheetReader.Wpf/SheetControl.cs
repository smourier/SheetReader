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

        public static readonly DependencyProperty RowHeightProperty = DependencyProperty.Register(nameof(RowHeight),
            typeof(double),
            typeof(SheetControl),
            new UIPropertyMetadata(20.0));

        public static readonly DependencyProperty GridLineSizeProperty = DependencyProperty.Register(nameof(GridLineSize),
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
            new UIPropertyMetadata(TextAlignment.Justify));

        public static readonly DependencyProperty TextTrimmingProperty = DependencyProperty.Register(nameof(TextTrimming),
            typeof(TextTrimming),
            typeof(SheetControl),
            new UIPropertyMetadata(TextTrimming.CharacterEllipsis));

        public BookDocumentSheet Sheet { get => (BookDocumentSheet)GetValue(SheetProperty); set => SetValue(SheetProperty, value); }
        public double ColumnWidth { get { return (double)GetValue(ColumnWidthProperty); } set { SetValue(ColumnWidthProperty, value); } }
        public double RowHeight { get { return (double)GetValue(RowHeightProperty); } set { SetValue(RowHeightProperty, value); } }
        public double GridLineSize { get => (double)GetValue(GridLineSizeProperty); set => SetValue(GridLineSizeProperty, value); }
        public Brush LineBrush { get { return (Brush)GetValue(LineBrushProperty); } set { SetValue(LineBrushProperty, value); } }
        public Brush HeaderBrush { get { return (Brush)GetValue(HeaderBrushProperty); } set { SetValue(HeaderBrushProperty, value); } }
        public TextAlignment TextAlignment { get { return (TextAlignment)GetValue(TextAlignmentProperty); } set { SetValue(TextAlignmentProperty, value); } }
        public TextTrimming TextTrimming { get { return (TextTrimming)GetValue(TextTrimmingProperty); } set { SetValue(TextTrimmingProperty, value); } }

        private double GetColWidth() => Math.Max(ColumnWidth, 10);
        private double GetRowHeight() => Math.Max(RowHeight, 10);
        private double GetRowMargin() => 60;

        static SheetControl()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(SheetControl), new FrameworkPropertyMetadata(typeof(SheetControl)));
        }

        private static void Log(object? message = null, [CallerMemberName] string? methodName = null) => EventProvider.Default.WriteMessageEvent(methodName + ":" + message);

        private ScrollViewer? _scrollViewer;
        private SheetGrid? _grid;
        private MovingColumn? _movingColumn;
        private readonly Dictionary<int, SheetControlColumn> _columnSettings = [];

        public SheetControl()
        {
            //Focusable = true;
        }

        public SheetControlHitTestResult HitTest(Point point) => _grid?.HitTest(point) ?? new();

        protected virtual void OnSheetChanged()
        {
            _columnSettings.Clear();
            var sheet = Sheet;
            if (sheet != null)
            {
                foreach (var kv in sheet.Columns)
                {
                    _columnSettings[kv.Key] = new SheetControlColumn(kv.Value) { Width = GetColWidth() };
                }
            }

            _grid?.InvalidateMeasure();
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
            Log("OnPreviewKeyDown:" + e.Key);
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
            Log("ReleaseMouseCapture commit:" + commit);
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
                Log("Moving col:" + result.MovingColumnIndex.Value);
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
            private double _width = column.Width;

            public SheetControlColumn Column { get; } = column;
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
                    var width = Math.Max(_current.X - Start.X, minSize);
                    if (width == _width)
                        return;

                    control._grid?.InvalidateVisual();
                }
            }
        }

        private sealed class SheetGrid(SheetControl control) : UIElement
        {
            private bool IsSheetVisible() => control.Sheet != null && control.Sheet.Columns.Count > 0 && control.Sheet.Rows.Count > 0 && control._scrollViewer != null;

            public SheetControlHitTestResult HitTest(Point point)
            {
                var result = new SheetControlHitTestResult();
                if (IsSheetVisible())
                {
                    var colWidth = control.GetColWidth();
                    var colFullWidth = colWidth + control.GridLineSize;
                    var rowHeight = control.GetRowHeight();
                    var rowFullHeight = rowHeight + control.GridLineSize;
                    var rowsHeaderWidth = control.GetRowMargin();
                    var rowsHeaderFullWidth = rowsHeaderWidth + control.GridLineSize;
                    var columnsHeaderHeight = rowHeight;
                    var columnsHeaderFullHeight = columnsHeaderHeight + control.GridLineSize;

                    var offsetX = control._scrollViewer!.HorizontalOffset;
                    var offsetY = control._scrollViewer.VerticalOffset;

                    var x = point.X + offsetX;
                    var y = point.Y + offsetY;

                    var columnIndex = (x - rowsHeaderFullWidth) / colFullWidth;
                    var rowIndex = (y - columnsHeaderFullHeight) / rowFullHeight;

                    result.RowCol = new RowCol { ColumnIndex = (int)Math.Floor(columnIndex), RowIndex = (int)Math.Floor(rowIndex) };
                    result.Cell = control.Sheet?.GetCell(result.RowCol);

                    // independent from scrollviewer
                    result.IsOverRowHeader = point.X >= 0 && point.X <= rowsHeaderFullWidth;
                    result.IsOverColumnHeader = point.Y >= 0 && point.Y <= columnsHeaderFullHeight;

                    if (!result.IsOverRowHeader)
                    {
                        var mod = (point.X + offsetX - rowsHeaderFullWidth) % colFullWidth;
                        const int tolerance = 4;
                        if (mod < tolerance)
                        {
                            result.MovingColumnIndex = result.RowCol.ColumnIndex - 1;
                        }
                        else if ((colFullWidth - mod) < tolerance)
                        {
                            result.MovingColumnIndex = result.RowCol.ColumnIndex;
                        }
                    }
                }
                return result;
            }

            protected override Size MeasureCore(Size availableSize)
            {
                if (!IsSheetVisible())
                    return new Size();

                var colWidth = control.GetColWidth();
                var colFullWidth = colWidth + control.GridLineSize;
                var rowHeight = control.GetRowHeight();
                var rowFullHeight = rowHeight + control.GridLineSize;
                var rowsHeaderWidth = control.GetRowMargin();
                var rowsHeaderFullWidth = rowsHeaderWidth + control.GridLineSize;
                var columnsHeaderHeight = rowHeight;
                var columnsHeaderFullHeight = columnsHeaderHeight + control.GridLineSize;
                return new Size(rowsHeaderFullWidth + colFullWidth * (control.Sheet.LastColumnIndex!.Value + 1), columnsHeaderFullHeight + rowFullHeight * (control.Sheet.LastRowIndex!.Value + 1));
            }

            protected override void OnRender(DrawingContext drawingContext)
            {
                if (!IsSheetVisible())
                    return;

                var viewWidth = control._scrollViewer!.ViewportWidth;
                if (viewWidth == 0)
                    return;

                var viewHeight = control._scrollViewer.ViewportHeight;
                if (viewHeight == 0)
                    return;

                var pen = new Pen(control.LineBrush, control.GridLineSize);
                var typeFace = new Typeface(control.FontFamily, control.FontStyle, control.FontWeight, control.FontStretch);

                var dpi = VisualTreeHelper.GetDpi(this);
                var colWidth = control.GetColWidth();
                var colFullWidth = colWidth + control.GridLineSize;
                var rowHeight = control.GetRowHeight();
                var rowFullHeight = rowHeight + control.GridLineSize;
                var rowsHeaderWidth = control.GetRowMargin();
                var rowsHeaderFullWidth = rowsHeaderWidth + control.GridLineSize;
                var columnsHeaderHeight = rowHeight;
                var columnsHeaderFullHeight = columnsHeaderHeight + control.GridLineSize;

                var offsetX = control._scrollViewer.HorizontalOffset;
                var offsetY = control._scrollViewer.VerticalOffset;
                //Log("offset:" + offsetX + " x " + offsetY + " view:" + viewWidth + " x " + viewHeight);

                var firstDrawnColumnIndex = Math.Max((int)((offsetX - rowsHeaderFullWidth) / colFullWidth), 0);
                var lastDrawnColumnIndex = Math.Max(Math.Min((int)((offsetX - rowsHeaderFullWidth + viewWidth) / colFullWidth), control.Sheet.LastColumnIndex!.Value), firstDrawnColumnIndex);

                var firstDrawnRowIndex = Math.Max((int)((offsetY - columnsHeaderFullHeight) / rowFullHeight), 0);
                var lastDrawnRowIndex = Math.Max(Math.Min((int)((offsetY - columnsHeaderFullHeight + viewHeight) / rowFullHeight), control.Sheet.LastRowIndex!.Value), firstDrawnRowIndex);

                //Log("col:" + firstDrawnColumnIndex + " => " + lastDrawnColumnIndex + " row:" + firstDrawnRowIndex + " => " + lastDrawnRowIndex);

                double yoffset;
                var cellsRect = new Rect(offsetX + rowsHeaderFullWidth, offsetY + columnsHeaderFullHeight, viewWidth - rowsHeaderFullWidth, viewHeight - columnsHeaderFullHeight);
                drawingContext.PushClip(new RectangleGeometry(cellsRect));
                for (var i = firstDrawnRowIndex; i <= lastDrawnRowIndex; i++)
                {
                    if (!control.Sheet.Rows.TryGetValue(i, out var row))
                        continue;

                    // draw cell
                    for (var j = firstDrawnColumnIndex; j <= lastDrawnColumnIndex; j++)
                    {
                        if (!row.Cells.TryGetValue(j, out var cell))
                            continue;

                        var text = string.Format(CultureInfo.CurrentCulture, "{0}", cell.Value);
                        if (string.IsNullOrEmpty(text))
                            continue;

                        var formattedCell = new FormattedText(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, typeFace, control.FontSize, control.Foreground, dpi.PixelsPerDip)
                        {
                            Trimming = control.TextTrimming,
                            TextAlignment = control.TextAlignment,
                            MaxTextWidth = colWidth,
                            MaxLineCount = 1
                        };

                        var y = columnsHeaderFullHeight + rowFullHeight * row.RowIndex;
                        var x = rowsHeaderFullWidth + colFullWidth * cell.ColumnIndex;
                        yoffset = (rowHeight - formattedCell.Height) / 2; // center vertically
                        drawingContext.DrawText(formattedCell, new Point(x + pen.Thickness, y + yoffset + pen.Thickness));
                        //Log("trace " + i + " x " + j + " " + text);
                    }
                }
                drawingContext.Pop();

                var maxColumnsWidth = (lastDrawnColumnIndex + 1) * colFullWidth + rowsHeaderFullWidth;
                var maxRowsHeight = (lastDrawnRowIndex + 1) * rowFullHeight + columnsHeaderFullHeight;

                // draw rows
                var rowsRect = new Rect(offsetX, offsetY + columnsHeaderHeight, viewWidth, Math.Min(viewHeight, maxRowsHeight) - columnsHeaderHeight);
                drawingContext.PushClip(new RectangleGeometry(rowsRect));
                rowsRect.Width = rowsHeaderWidth;
                drawingContext.DrawRectangle(control.HeaderBrush, null, rowsRect);
                for (var i = firstDrawnRowIndex; i <= lastDrawnRowIndex + 1; i++)
                {
                    var y = columnsHeaderFullHeight + rowFullHeight * i;
                    drawingContext.DrawLine(pen, new Point(offsetX, y), new Point(offsetX + Math.Min(maxColumnsWidth, viewWidth), y));
                    if (i > lastDrawnRowIndex)
                        break;

                    // draw row name
                    var name = (i + 1).ToString();
                    var formattedRow = new FormattedText(name, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, typeFace, control.FontSize, control.Foreground, dpi.PixelsPerDip)
                    {
                        MaxTextWidth = rowsHeaderWidth,
                        MaxLineCount = 1
                    };

                    var xoffset = offsetX + (rowsHeaderWidth - formattedRow.Width) / 2; // center horizontally
                    yoffset = (rowHeight - formattedRow.Height) / 2; // center vertically
                    drawingContext.DrawText(formattedRow, new Point(xoffset + pen.Thickness, columnsHeaderHeight + rowFullHeight * i + yoffset + pen.Thickness));
                }
                drawingContext.Pop();

                // draw columns
                var columnsRect = new Rect(offsetX + rowsHeaderWidth, offsetY, Math.Min(viewWidth, maxColumnsWidth) - rowsHeaderWidth, viewHeight);
                drawingContext.PushClip(new RectangleGeometry(columnsRect));
                columnsRect.Height = columnsHeaderHeight;
                drawingContext.DrawRectangle(control.HeaderBrush, null, columnsRect);
                for (var i = firstDrawnColumnIndex; i <= lastDrawnColumnIndex + 1; i++)
                {
                    var x = rowsHeaderFullWidth + colFullWidth * i;
                    drawingContext.DrawLine(pen, new Point(x, offsetY), new Point(x, offsetY + Math.Min(maxRowsHeight, viewHeight)));
                    if (i > lastDrawnColumnIndex)
                        break;

                    // draw col name
                    var name = Row.GetExcelColumnName(i);
                    var formattedCol = new FormattedText(name, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, typeFace, control.FontSize, control.Foreground, dpi.PixelsPerDip)
                    {
                        MaxTextWidth = colWidth,
                        MaxLineCount = 1
                    };

                    var xoffset = (colWidth - formattedCol.Width) / 2; // center horizontally
                    yoffset = offsetY + (columnsHeaderHeight - formattedCol.Height) / 2; // center vertically
                    drawingContext.DrawText(formattedCol, new Point(xoffset + rowsHeaderFullWidth + i * colWidth + pen.Thickness, yoffset + pen.Thickness));
                }
                drawingContext.Pop();
            }
        }
    }
}
