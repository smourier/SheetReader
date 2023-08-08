using SheetReader.Utilities;

namespace SheetReader
{
    public class Column
    {
        public virtual string? Name { get; set; }
        public virtual int Index { get; set; }

        public override string ToString() => Name.Nullify() ?? Index.ToString();
    }
}
