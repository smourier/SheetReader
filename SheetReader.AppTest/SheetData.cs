using System.Collections.Generic;

namespace SheetReader.AppTest
{
    public class SheetData
    {
        public string? Name { get; set; }
        public bool IsHidden { get; set; }
        public List<RowData> Rows { get; } = new List<RowData>();
        public List<Column>? Columns { get; set; }
        public int LastColumnIndex { get; set; }
        public int LastRowIndex { get; set; }

        public override string ToString() => Name ?? string.Empty;
    }

    public class RowData
    {
        public RowData(int index, List<Cell> cells)
        {
            RowIndex = index;
            Cells = cells;
        }

        public int RowIndex { get; }
        public List<Cell> Cells { get; }
    }
}
