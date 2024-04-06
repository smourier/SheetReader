using System;
using System.Collections.Generic;

namespace SheetReader
{
    public class BookDocumentSheet
    {
        private readonly Dictionary<int, BookDocumentRow> _rows = [];
        private readonly Dictionary<int, Column> _columns = [];

        public BookDocumentSheet(Sheet sheet)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            Name = sheet.Name ?? string.Empty;
            IsHidden = !sheet.IsVisible;

            foreach (var row in sheet.EnumerateRows())
            {
                var rowData = CreateRow(row);
                _rows[row.Index] = rowData;

                if (LastRowIndex == null || row.Index > LastRowIndex)
                {
                    LastRowIndex = row.Index;
                }

                if (FirstRowIndex == null || row.Index < FirstRowIndex)
                {
                    FirstRowIndex = row.Index;
                }
            }

            foreach (var col in sheet.EnumerateColumns())
            {
                _columns[col.Index] = col;
                if (LastColumnIndex == null || col.Index > LastColumnIndex)
                {
                    LastColumnIndex = col.Index;
                }

                if (FirstColumnIndex == null || col.Index < FirstColumnIndex)
                {
                    FirstColumnIndex = col.Index;
                }
            }
        }

        public string Name { get; }
        public bool IsHidden { get; }
        public int? FirstColumnIndex { get; }
        public int? LastColumnIndex { get; }
        public int? FirstRowIndex { get; }
        public int? LastRowIndex { get; }
        public IReadOnlyDictionary<int, BookDocumentRow> Rows => _rows;
        public IReadOnlyDictionary<int, Column> Columns => _columns;

        public override string ToString() => Name;

        public Cell? GetCell(RowCol? rowCol)
        {
            if (rowCol == null)
                return null;

            return GetCell(rowCol.RowIndex, rowCol.ColumnIndex);
        }

        public virtual Cell? GetCell(int rowIndex, int columnIndex)
        {
            if (!_rows.TryGetValue(rowIndex, out var row))
                return null;

            row.Cells.TryGetValue(columnIndex, out var cell);
            return cell;
        }

        protected virtual BookDocumentRow CreateRow(Row row) => new(row);
    }
}
