using SheetReader.Utilities;

namespace SheetReader
{
    public class Column : IWithValue
    {
        public virtual int Index { get; set; }
        public virtual string? Name { get; set; }
        object? IWithValue.Value => Name;

        public override string ToString() => Extensions.Nullify(Name) ?? Index.ToString();
    }
}
