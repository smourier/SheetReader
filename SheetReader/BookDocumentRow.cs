using System;
using System.Collections.Generic;

namespace SheetReader
{
    public class BookDocumentRow
    {
        private IDictionary<int, BookDocumentCell> _cells;

        public BookDocumentRow(BookDocument book, BookDocumentSheet sheet, Row row)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(book);
            ArgumentNullException.ThrowIfNull(row);
            _cells = CreateCells();
            if (_cells == null)
                throw new InvalidOperationException();

            Row = row;
            RowIndex = row.Index;
            SortIndex = RowIndex;
            IsHidden = !row.IsVisible;
            foreach (var cell in row.EnumerateCells())
            {
                var bdCell = CreateCell(cell);
                if (bdCell == null)
                    continue;

                if (!sheet.EnsureColumn(book, cell.ColumnIndex))
                    break;

                _cells[cell.ColumnIndex] = bdCell;

                if (LastCellIndex == null || row.Index > LastCellIndex)
                {
                    LastCellIndex = row.Index;
                }

                if (FirstCellIndex == null || row.Index < FirstCellIndex)
                {
                    FirstCellIndex = row.Index;
                }

                var e = new StateChangedEventArgs(StateChangedType.CellAdded, sheet, this, null, bdCell);
                book.OnStateChanged(this, e);
                if (e.Cancel)
                    break;
            }
        }

        public Row Row { get; }
        public int RowIndex { get; } // orginal index
        public virtual int SortIndex { get; set; } // current sorted index
        public virtual bool IsHidden { get; }
        public int? FirstCellIndex { get; }
        public int? LastCellIndex { get; }
        public IDictionary<int, BookDocumentCell> Cells => _cells;

        protected virtual internal IDictionary<int, BookDocumentCell> CreateCells() => new Dictionary<int, BookDocumentCell>();
        protected virtual BookDocumentCell CreateCell(Cell cell)
        {
            ArgumentNullException.ThrowIfNull(cell);
            if (cell is Book.JsonCell json)
                return cell.IsError ? new BookDocumentJsonCellError(json) : new BookDocumentJsonCell(json);

            return cell.IsError ? new BookDocumentCellError(cell) : new BookDocumentCell(cell);
        }

        protected virtual internal void ReplaceCells(IDictionary<int, BookDocumentCell> cells)
        {
            ArgumentNullException.ThrowIfNull(cells);
            _cells = cells;
        }

        public override string ToString() => RowIndex.ToString();
    }
}
