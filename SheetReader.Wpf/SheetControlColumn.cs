using System;

namespace SheetReader.Wpf
{
    public class SheetControlColumn
    {
        public SheetControlColumn(Column column)
        {
            ArgumentNullException.ThrowIfNull(column);
            Column = column;
        }

        public Column Column { get; }
        public virtual double Width { get; set; }

        public override string ToString() => $"{Column} {Width}";
    }
}
