using System.Collections.Generic;

namespace SheetReader
{
    public abstract class Row
    {
        public virtual int Index { get; set; }
        public abstract IEnumerable<Cell> EnumerateCells();

        public override string ToString() => Index.ToString();
    }
}
