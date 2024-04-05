using System;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
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

        public BookDocumentSheet Sheet { get => (BookDocumentSheet)GetValue(SheetProperty); set => SetValue(SheetProperty, value); }
        public double ColumnWidth { get { return (double)GetValue(ColumnWidthProperty); } set { SetValue(ColumnWidthProperty, value); } }
        public double RowHeight { get { return (double)GetValue(RowHeightProperty); } set { SetValue(RowHeightProperty, value); } }
        public double GridLineSize { get => (double)GetValue(GridLineSizeProperty); set => SetValue(GridLineSizeProperty, value); }

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

        protected virtual void OnSheetChanged()
        {
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

        private void OnScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            _grid?.InvalidateVisual();
        }

        private static string GetExcelColumnName(int index)
        {
            index++;
            var name = string.Empty;
            while (index > 0)
            {
                var mod = (index - 1) % 26;
                name = (char)('A' + mod) + name;
                index = (index - mod) / 26;
            }
            return name;
        }

        private sealed class SheetGrid(SheetControl control) : UIElement
        {
            protected override Size MeasureCore(Size availableSize)
            {
                if (control.Sheet == null)
                    return new Size();

                var rowHeight = control.GetRowHeight();
                var rowMargin = control.GetRowMargin();
                var headerHeight = rowHeight;
                return new Size(rowMargin + control.GetColWidth() * (control.Sheet.LastColumnIndex + 1), headerHeight + rowHeight * (control.Sheet.LastRowIndex + 1));
            }

            protected override void OnRender(DrawingContext drawingContext)
            {
                if (control.Sheet == null || control.Sheet.Columns.Count == 0 || control.Sheet.Rows.Count == 0)
                    return;

                var pen = new Pen(Brushes.Gray, control.GridLineSize);
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

                var offsetX = control._scrollViewer?.HorizontalOffset ?? 0;
                var offsetY = control._scrollViewer?.VerticalOffset ?? 0;
                var viewWidth = control._scrollViewer?.ViewportWidth ?? RenderSize.Width;
                var viewHeight = control._scrollViewer?.ViewportHeight ?? RenderSize.Height;

                var firstDrawnColumnIndex = Math.Max((int)((offsetX - rowsHeaderFullWidth) / colFullWidth), 0);
                var lastDrawnColumnIndex = Math.Max(Math.Min((int)((offsetX - rowsHeaderFullWidth + viewWidth) / colFullWidth), control.Sheet.LastColumnIndex), firstDrawnColumnIndex);

                var firstDrawnRowIndex = Math.Max((int)((offsetY - columnsHeaderFullHeight) / rowFullHeight), 0);
                var lastDrawnRowIndex = Math.Max(Math.Min((int)((offsetY - columnsHeaderFullHeight + viewHeight) / rowFullHeight), control.Sheet.LastRowIndex), firstDrawnRowIndex);

                Log("firstCol:" + firstDrawnColumnIndex + " lastCol:" + lastDrawnColumnIndex + " firstRow:" + firstDrawnRowIndex + " lastRow:" + lastDrawnRowIndex);

                // header backgrounds
                drawingContext.DrawRectangle(Brushes.LightGray, null, new Rect(offsetX, columnsHeaderHeight, rowsHeaderWidth, (control.Sheet.LastRowIndex + 1) * rowHeight));
                drawingContext.DrawRectangle(Brushes.LightGray, null, new Rect(rowsHeaderWidth, offsetY, (control.Sheet.LastColumnIndex + 1) * colWidth, columnsHeaderHeight));

                var maxWidth = Math.Min(colWidth * (control.Sheet.LastColumnIndex + 1) + rowsHeaderWidth, RenderSize.Width);
                var maxHeight = Math.Min(rowHeight * (control.Sheet.LastRowIndex + 2), RenderSize.Height);

                // includes col header
                for (var i = 0; i < control.Sheet.LastColumnIndex + 2 + 1; i++)
                {
                    double x;
                    if (i == 0)
                    {
                        x = 0;
                    }
                    else
                    {
                        x = colWidth * i - (colWidth - rowsHeaderWidth);
                    }

                    drawingContext.DrawLine(pen, new Point(x, offsetY), new Point(x, maxHeight));

                    // draw col name
                    if (i < control.Sheet.LastColumnIndex + 1)
                    {
                        var name = GetExcelColumnName(i);
                        var formatted = new FormattedText(name, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, typeFace, control.FontSize, control.Foreground, dpi.PixelsPerDip)
                        {
                            MaxTextWidth = rowsHeaderWidth,
                            MaxLineCount = 1
                        };

                        var xoffset = (colWidth - formatted.Width) / 2; // center horizontally
                        var yoffset = offsetY + (columnsHeaderHeight - formatted.Height) / 2; // center vertically
                        drawingContext.DrawText(formatted, new Point(xoffset + rowsHeaderWidth + i * colWidth + pen.Thickness, yoffset + pen.Thickness));
                    }
                }

                // includess row margin
                for (var i = 0; i < control.Sheet.LastRowIndex + 2 + 1; i++)
                {
                    var y = rowHeight * i;
                    drawingContext.DrawLine(pen, new Point(offsetX, y), new Point(maxWidth, y));

                    // draw row #
                    if (i < control.Sheet.LastRowIndex + 1)
                    {
                        var formatted = new FormattedText((i + 1).ToString(), CultureInfo.CurrentCulture, FlowDirection.LeftToRight, typeFace, control.FontSize, control.Foreground, dpi.PixelsPerDip)
                        {
                            MaxTextWidth = rowsHeaderWidth,
                            MaxLineCount = 1
                        };

                        var xoffset = offsetX + (rowsHeaderWidth - formatted.Width) / 2; // center horizontally
                        var yoffset = (rowHeight - formatted.Height) / 2; // center vertically
                        drawingContext.DrawText(formatted, new Point(xoffset + pen.Thickness, columnsHeaderHeight + rowHeight * i + yoffset + pen.Thickness));
                    }
                }

                for (var i = firstDrawnRowIndex; i <= lastDrawnRowIndex; i++)
                {
                    if (!control.Sheet.Rows.TryGetValue(i, out var row))
                        continue;

                    for (var j = firstDrawnColumnIndex; j <= lastDrawnColumnIndex; j++)
                    {
                        if (j >= row.Cells.Count)
                            continue;

                        var cell = row.Cells[j];
                        var text = string.Format("{0}", cell.Value);
                        var formatted = new FormattedText(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, typeFace, control.FontSize, control.Foreground, dpi.PixelsPerDip)
                        {
                            MaxTextWidth = colWidth,
                            MaxLineCount = 1
                        };

                        var y = columnsHeaderHeight + rowHeight * row.RowIndex;
                        var x = rowsHeaderWidth + colWidth * cell.ColumnIndex;
                        var yoffset = (rowHeight - formatted.Height) / 2; // center vertically
                        drawingContext.DrawText(formatted, new Point(x + pen.Thickness, y + yoffset + pen.Thickness));
                    }
                }
            }
        }
    }
}
