using System.Windows;
using System.Windows.Media;

namespace SheetReader.Wpf
{
    public class GetStyleContext
    {
        public virtual RowCol? RowCol { get; set; }
        public virtual double PixelsPerDip { get; set; }
        public virtual int? MaxLineCount { get; set; }
        public virtual double? ColumnWidth { get; set; }
        public virtual Typeface? Typeface { get; set; }
        public virtual double FontSize { get; set; }
        public virtual Brush? Foreground { get; set; }
        public virtual Brush? ErrorForeground { get; set; }
        public virtual TextTrimming? TextTrimming { get; set; }
        public virtual TextAlignment? TextAlignment { get; set; }
    }
}
