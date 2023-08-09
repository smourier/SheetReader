using SheetReader.Utilities;

namespace SheetReader
{
    public class Column
    {
        public virtual int Index { get; set; }
        public virtual string? Name { get; set; }

        public override string ToString() => Name.Nullify() ?? Index.ToString();
    }
}
