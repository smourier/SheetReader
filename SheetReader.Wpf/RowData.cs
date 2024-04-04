using System.Collections.Generic;

namespace SheetReader.Wpf
{
    public class RowData(int index, List<Cell> cells)
    {
        public int RowIndex { get; } = index;
        public List<Cell> Cells { get; } = cells;
    }
}
