using System;
using System.Globalization;
using System.Windows;
using System.Windows.Media;

namespace SheetReader.AppTest
{
    public class SheetControl : FrameworkElement
    {
        public static readonly DependencyProperty SheetProperty = DependencyProperty.Register(nameof(Sheet), typeof(SheetData), typeof(SheetControl),
            new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsMeasure | FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty ColumnSizeProperty = DependencyProperty.Register(nameof(ColumnSize), typeof(double), typeof(SheetControl),
            new FrameworkPropertyMetadata(100.0, FrameworkPropertyMetadataOptions.AffectsMeasure | FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty RowHeightProperty = DependencyProperty.Register(nameof(RowHeight), typeof(double), typeof(SheetControl),
            new FrameworkPropertyMetadata(20.0, FrameworkPropertyMetadataOptions.AffectsMeasure | FrameworkPropertyMetadataOptions.AffectsRender));

        public SheetData Sheet { get { return (SheetData)GetValue(SheetProperty); } set { SetValue(SheetProperty, value); } }
        public double ColumnSize { get { return (double)GetValue(ColumnSizeProperty); } set { SetValue(ColumnSizeProperty, value); } }
        public double RowHeight { get { return (double)GetValue(RowHeightProperty); } set { SetValue(RowHeightProperty, value); } }

        private static readonly Typeface _typeFace = new Typeface("Lucida Console");
        private static readonly Pen _pen = new(Brushes.Gray, 1);

        private double GetColSize() => Math.Max(ColumnSize, 10);
        private double GetRowHeight() => Math.Max(RowHeight, 10);
        private double GetRowMargin() => 60;

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

        protected override Size MeasureOverride(Size availableSize)
        {
            if (Sheet == null)
                return new Size();

            var rowHeight = GetRowHeight();
            var rowMargin = GetRowMargin();
            var headerHeight = rowHeight;
            return new Size(rowMargin + GetColSize() * (Sheet.LastColumnIndex + 1), headerHeight + rowHeight * (Sheet.LastRowIndex + 1));
        }

        protected override void OnRender(DrawingContext drawingContext)
        {
            if (Sheet == null)
                return;

            var dpi = VisualTreeHelper.GetDpi(this);
            var colSize = GetColSize();
            var rowHeight = GetRowHeight();
            var fontSize = rowHeight * 0.6;

            var rowMargin = GetRowMargin();
            var headerHeight = rowHeight;

            // header backgrounds
            drawingContext.DrawRectangle(Brushes.LightGray, null, new Rect(0, headerHeight, rowMargin, (Sheet.LastRowIndex + 1) * rowHeight));
            drawingContext.DrawRectangle(Brushes.LightGray, null, new Rect(rowMargin, 0, (Sheet.LastColumnIndex + 1) * colSize, headerHeight));

            // includes col header
            for (var i = 0; i < Sheet.LastColumnIndex + 2 + 1; i++)
            {
                double x;
                if (i == 0)
                {
                    x = 0;
                }
                else
                {
                    x = colSize * i - (colSize - rowMargin);
                }
                drawingContext.DrawLine(_pen, new Point(x, 0), new Point(x, RenderSize.Height));

                // draw col name
                if (i < Sheet.LastColumnIndex + 1)
                {
                    var name = GetExcelColumnName(i);
                    var formatted = new FormattedText(name, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, _typeFace, fontSize, Brushes.Black, dpi.PixelsPerDip)
                    {
                        MaxTextWidth = rowMargin,
                        MaxLineCount = 1
                    };

                    var xoffset = (colSize - formatted.Width) / 2; // center horizontally
                    var yoffset = (headerHeight - formatted.Height) / 2; // center vertically
                    drawingContext.DrawText(formatted, new Point(xoffset + rowMargin + i * colSize + _pen.Thickness, yoffset + _pen.Thickness));
                }
            }

            // includess row margin
            for (var i = 0; i < Sheet.LastRowIndex + 2 + 1; i++)
            {
                var y = rowHeight * i;
                drawingContext.DrawLine(_pen, new Point(0, y), new Point(RenderSize.Width, y));

                // draw row #
                if (i < Sheet.LastRowIndex + 1)
                {
                    var formatted = new FormattedText((i + 1).ToString(), CultureInfo.CurrentCulture, FlowDirection.LeftToRight, _typeFace, fontSize, Brushes.Black, dpi.PixelsPerDip)
                    {
                        MaxTextWidth = rowMargin,
                        MaxLineCount = 1
                    };

                    var xoffset = (rowMargin - formatted.Width) / 2; // center horizontally
                    var yoffset = (rowHeight - formatted.Height) / 2; // center vertically
                    drawingContext.DrawText(formatted, new Point(xoffset + _pen.Thickness, headerHeight + rowHeight * i + yoffset + _pen.Thickness));
                }
            }

            foreach (var row in Sheet.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    var text = string.Format("{0}", cell.Value);
                    var formatted = new FormattedText(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, _typeFace, fontSize, Brushes.Black, dpi.PixelsPerDip)
                    {
                        MaxTextWidth = colSize,
                        MaxLineCount = 1
                    };

                    var y = headerHeight + rowHeight * row.RowIndex;
                    var x = rowMargin + colSize * cell.ColumnIndex;
                    var yoffset = (rowHeight - formatted.Height) / 2; // center vertically
                    drawingContext.DrawText(formatted, new Point(x + _pen.Thickness, y + yoffset + _pen.Thickness));
                }
            }
        }
    }
}
