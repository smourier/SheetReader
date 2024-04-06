using System;
using System.Collections.Generic;

namespace SheetReader
{
    public class BookDocumentRow
    {
        private readonly Dictionary<int, Cell> _cells = [];

        public BookDocumentRow(Row row)
        {
            ArgumentNullException.ThrowIfNull(row);
            RowIndex = row.Index;
            IsHidden = !row.IsVisible;
            foreach (var cell in row.EnumerateCells())
            {
                _cells[cell.ColumnIndex] = cell;

                if (LastCellIndex == null || row.Index > LastCellIndex)
                {
                    LastCellIndex = row.Index;
                }

                if (FirstCellIndex == null || row.Index < FirstCellIndex)
                {
                    FirstCellIndex = row.Index;
                }
            }
        }

        public int RowIndex { get; }
        public bool IsHidden { get; }
        public int? FirstCellIndex { get; }
        public int? LastCellIndex { get; }
        public IReadOnlyDictionary<int, Cell> Cells => _cells;

        public override string ToString() => RowIndex.ToString();
    }
}
