using System;
using System.Collections.Generic;

namespace SheetReader
{
    public class BookDocumentRow
    {
        private readonly IDictionary<int, BookDocumentCell> _cells;

        public BookDocumentRow(Row row)
        {
            ArgumentNullException.ThrowIfNull(row);
            _cells = CreateCells();
            if (_cells == null)
                throw new InvalidOperationException();

            RowIndex = row.Index;
            IsHidden = !row.IsVisible;
            foreach (var cell in row.EnumerateCells())
            {
                var bdCell = CreateCell(cell);
                if (bdCell == null)
                    continue;

                _cells[cell.ColumnIndex] = CreateCell(cell);

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
        public IDictionary<int, BookDocumentCell> Cells => _cells;

        protected virtual IDictionary<int, BookDocumentCell> CreateCells() => new Dictionary<int, BookDocumentCell>();
        protected virtual BookDocumentCell CreateCell(Cell cell)
        {
            ArgumentNullException.ThrowIfNull(cell);
            return cell.IsError ? new BookDocumentCellError(cell) : new BookDocumentCell(cell);
        }

        public override string ToString() => RowIndex.ToString();
    }
}
