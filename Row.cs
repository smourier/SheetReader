using System.Collections.Generic;

namespace SheetReader
{
    public abstract class Row
    {
        public virtual int Index { get; set; }
        public virtual bool IsVisible { get; set; } = true;

        public abstract IEnumerable<Cell> EnumerateCells();

        public static string GetExcelColumnName(int index)
        {
            index++;
            var name = string.Empty;
            while (index > 0)
            {
                var mod = (index - 1) % 26;
                name = (char)('A' + mod) + name;
                index = (index - mod) / 26;
            }
            return name;
        }

        public override string ToString() => Index.ToString();
    }
}
