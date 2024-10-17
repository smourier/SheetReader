using System;

namespace SheetReader.Wpf
{
    public class SheetControlColumn
    {
        public SheetControlColumn(BookDocumentColumn column)
        {
            ArgumentNullException.ThrowIfNull(column);
            Column = column;
        }

        public BookDocumentColumn Column { get; }
        public virtual double Width { get; set; }

        public override string ToString() => $"{Column} {Width}";
    }
}
