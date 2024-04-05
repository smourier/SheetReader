using System.Collections.Generic;

namespace SheetReader
{
    public sealed class BookDocumentSheet
    {
        private readonly Dictionary<int, BookDocumentRow> _rows = [];
        private readonly Dictionary<int, Column> _columns = [];

        internal BookDocumentSheet(Sheet sheet)
        {
            Name = sheet.Name ?? string.Empty;
            IsHidden = !sheet.IsVisible;

            foreach (var row in sheet.EnumerateRows())
            {
                var rowData = new BookDocumentRow(row);
                _rows[row.Index] = rowData;

                if (row.Index > LastRowIndex)
                {
                    LastRowIndex = row.Index;
                }
            }

            foreach (var col in sheet.EnumerateColumns())
            {
                _columns[col.Index] = col;
                if (col.Index > LastColumnIndex)
                {
                    LastColumnIndex = col.Index;
                }
            }
        }

        public string Name { get; }
        public bool IsHidden { get; }
        public int LastColumnIndex { get; }
        public int LastRowIndex { get; }
        public IReadOnlyDictionary<int, BookDocumentRow> Rows => _rows;
        public IReadOnlyDictionary<int, Column> Columns => _columns;

        public override string ToString() => Name;
    }
}
