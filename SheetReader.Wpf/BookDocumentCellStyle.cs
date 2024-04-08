using System;
using System.Globalization;
using System.Windows;
using System.Windows.Media;
namespace SheetReader.Wpf
{
    public class BookDocumentCellStyle
    {
        public static Brush DefaultErrorForeground { get; set; } = Brushes.Red;
        public static BookDocumentCellStyle Empty { get; } = new();

        public virtual Typeface? Typeface { get; set; }
        public virtual double? FontSize { get; set; }
        public virtual Brush? Foreground { get; set; }
        public virtual TextTrimming? TextTrimming { get; set; }
        public virtual Brush? ErrorForeground { get; set; }
        public virtual TextAlignment? TextAlignment { get; set; }

        public virtual string? FormatCell(BookDocumentCell cell)
        {
            ArgumentNullException.ThrowIfNull(cell);
            return string.Format(CultureInfo.CurrentCulture, "{0}", cell.Value);
        }

        public virtual FormattedText? CreateCellFormattedText(StyleContext context, BookDocumentCell cell)
        {
            ArgumentNullException.ThrowIfNull(context);
            ArgumentNullException.ThrowIfNull(cell);
            if (cell.Value == null)
                return null;

            var text = FormatCell(cell);
            if (string.IsNullOrEmpty(text))
                return null;

            var fg = cell.IsError ? ErrorForeground ?? context.ErrorForeground ?? DefaultErrorForeground : Foreground ?? context.Foreground;
            var formatted = new FormattedText(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, Typeface ?? context.Typeface, FontSize ?? context.FontSize, fg, context.PixelsPerDip);
            if (context.CellWidth.HasValue)
            {
                formatted.MaxTextWidth = context.CellWidth.Value;
            }
            else if (context.ColumnWidth.HasValue)
            {
                formatted.MaxTextWidth = context.ColumnWidth.Value;
            }

            if (context.CellHeight.HasValue)
            {
                formatted.MaxTextHeight = context.CellHeight.Value;
            }
            else if (context.RowHeight.HasValue)
            {
                formatted.MaxTextHeight = context.RowHeight.Value;
            }

            if (context.MaxLineCount.HasValue)
            {
                formatted.MaxLineCount = context.MaxLineCount.Value;
            }

            if (TextTrimming.HasValue)
            {
                formatted.Trimming = TextTrimming.Value;
            }
            else if (context.TextTrimming.HasValue)
            {
                formatted.Trimming = context.TextTrimming.Value;
            }

            if (TextAlignment.HasValue)
            {
                formatted.TextAlignment = TextAlignment.Value;
            }
            else if (cell.IsError)
            {
                formatted.TextAlignment = System.Windows.TextAlignment.Center;
            }
            else if (context.TextAlignment.HasValue)
            {
                formatted.TextAlignment = context.TextAlignment.Value;
            }
            return formatted;
        }
    }
}
