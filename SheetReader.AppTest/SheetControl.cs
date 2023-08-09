using System;
using System.Globalization;
using System.Windows;
using System.Windows.Media;

namespace SheetReader.AppTest
{
    public class SheetControl : FrameworkElement
    {
        public static readonly DependencyProperty SheetProperty = DependencyProperty.Register(nameof(Sheet), typeof(SheetData), typeof(SheetControl),
            new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsMeasure | FrameworkPropertyMetadataOptions.AffectsArrange | FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty ColumnSizeProperty = DependencyProperty.Register(nameof(ColumnSize), typeof(double), typeof(SheetControl),
            new FrameworkPropertyMetadata(100.0, FrameworkPropertyMetadataOptions.AffectsMeasure | FrameworkPropertyMetadataOptions.AffectsArrange | FrameworkPropertyMetadataOptions.AffectsRender));

        public static readonly DependencyProperty RowHeightProperty = DependencyProperty.Register(nameof(RowHeight), typeof(double), typeof(SheetControl),
            new FrameworkPropertyMetadata(20.0, FrameworkPropertyMetadataOptions.AffectsMeasure | FrameworkPropertyMetadataOptions.AffectsArrange | FrameworkPropertyMetadataOptions.AffectsRender));

        public SheetData Sheet { get { return (SheetData)GetValue(SheetProperty); } set { SetValue(SheetProperty, value); } }
        public double ColumnSize { get { return (double)GetValue(ColumnSizeProperty); } set { SetValue(ColumnSizeProperty, value); } }
        public double RowHeight { get { return (double)GetValue(RowHeightProperty); } set { SetValue(RowHeightProperty, value); } }

        private static readonly Typeface _typeFace = new Typeface("Lucida Console");

        protected override Size MeasureOverride(Size availableSize)
        {
            return availableSize;
        }

        protected override Size ArrangeOverride(Size finalSize)
        {
            return finalSize;
        }

        protected override void OnRender(DrawingContext drawingContext)
        {
            if (Sheet == null)
                return;

            var dpi = VisualTreeHelper.GetDpi(this);
            var colSize = Math.Max(ColumnSize, 10);
            var rowHeight = Math.Max(RowHeight, 10);
            var fontSize = rowHeight * 0.6;

            var y = 0.0;
            foreach (var row in Sheet.Rows)
            {
                var x = 0.0;
                foreach (var cell in row)
                {
                    var text = string.Format("{0}", cell.Value);
                    var formatted = new FormattedText(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, _typeFace, fontSize, Brushes.Black, dpi.PixelsPerDip);
                    formatted.MaxTextWidth = colSize;
                    formatted.MaxLineCount = 1;
                    drawingContext.DrawText(formatted, new Point(x, y));
                    x += colSize;
                }
                y += rowHeight;
            }
        }
    }
}
