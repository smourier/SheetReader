using System.Collections.Generic;

namespace SheetReader
{
    public abstract class Sheet
    {
        public virtual string? Name { get; set; }
        public virtual bool IsVisible { get; set; } = true;

        public abstract IEnumerable<Column> EnumerateColumns();
        public abstract IEnumerable<Row> EnumerateRows();

        public override string ToString() => Name ?? string.Empty;
    }
}
