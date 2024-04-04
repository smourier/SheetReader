using System.Collections.Generic;

namespace SheetReader.Wpf
{
    public class SheetData
    {
        public string? Name { get; set; }
        public bool IsHidden { get; set; }
        public List<RowData> Rows { get; } = [];
        public List<Column> Columns { get; set; } = [];
        public int LastColumnIndex { get; set; }
        public int LastRowIndex { get; set; }

        public override string ToString() => Name ?? string.Empty;
    }
}
