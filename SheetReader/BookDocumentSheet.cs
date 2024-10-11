using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace SheetReader
{
    public class BookDocumentSheet
    {
        private readonly IDictionary<int, BookDocumentRow> _rows;
        private readonly IDictionary<int, Column> _columns;

        public BookDocumentSheet(BookDocument book, Sheet sheet)
        {
            ArgumentNullException.ThrowIfNull(book);
            ArgumentNullException.ThrowIfNull(sheet);
            _rows = CreateRows();
            _columns = CreateColumns();
            if (_rows == null || _columns == null)
                throw new InvalidOperationException();

            Name = sheet.Name ?? string.Empty;
            IsHidden = !sheet.IsVisible;

            var e = new StateChangedEventArgs(StateChangedType.SheetAdded, this);
            book.OnStateChanged(this, e);
            if (e.Cancel)
                return;
        }

        public virtual void Load(BookDocument book, Sheet sheet)
        {
            ArgumentNullException.ThrowIfNull(book);
            ArgumentNullException.ThrowIfNull(sheet);

            StateChangedEventArgs e;
            foreach (var row in sheet.EnumerateRows())
            {
                var rowData = CreateRow(book, row);
                if (rowData == null)
                    continue;

                _rows[row.Index] = rowData;

                if (LastRowIndex == null || row.Index > LastRowIndex)
                {
                    LastRowIndex = row.Index;
                }

                if (FirstRowIndex == null || row.Index < FirstRowIndex)
                {
                    FirstRowIndex = row.Index;
                }

                e = new StateChangedEventArgs(StateChangedType.RowAdded, this, rowData);
                book.OnStateChanged(this, e);
                if (e.Cancel)
                    break;
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

                e = new StateChangedEventArgs(StateChangedType.ColumnAddded, this, null, col);
                book.OnStateChanged(this, e);
                if (e.Cancel)
                    break;
            }

            if (_columns.Count == 0 && _rows.Count > 0)
            {
                ColumnsHaveBeenGenerated = true;
                for (var i = 0; i < _rows[0].Cells.Count; i++)
                {
                    var col = new Column { Index = i };
                    _columns[i] = col;
                    if (!LastColumnIndex.HasValue || col.Index > LastColumnIndex)
                    {
                        LastColumnIndex = col.Index;
                    }

                    if (!FirstColumnIndex.HasValue || col.Index < FirstColumnIndex)
                    {
                        FirstColumnIndex = col.Index;
                    }

                    e = new StateChangedEventArgs(StateChangedType.ColumnAddded, this, null, col);
                    book.OnStateChanged(this, e);
                    if (e.Cancel)
                        break;
                }
            }
        }

        public virtual string Name { get; }
        public virtual bool IsHidden { get; }
        public bool ColumnsHaveBeenGenerated { get; protected set; }
        public int? FirstColumnIndex { get; protected set; }
        public int? LastColumnIndex { get; protected set; }
        public int? FirstRowIndex { get; protected set; }
        public int? LastRowIndex { get; protected set; }
        public IDictionary<int, BookDocumentRow> Rows => _rows;
        public IDictionary<int, Column> Columns => _columns;

        public override string ToString() => Name;

        public BookDocumentCell? GetCell(RowCol? rowCol)
        {
            if (rowCol == null)
                return null;

            return GetCell(rowCol.RowIndex, rowCol.ColumnIndex);
        }

        public virtual BookDocumentCell? GetCell(int rowIndex, int columnIndex)
        {
            if (!_rows.TryGetValue(rowIndex, out var row))
                return null;

            row.Cells.TryGetValue(columnIndex, out var cell);
            return cell;
        }

        public virtual string? FormatValue(object? value)
        {
            if (value is null)
                return null;

            if (value is string s)
                return s;

            if (value is IDictionary dictionary)
            {
                var sb = new StringBuilder("{");
                var first = true;
                foreach (DictionaryEntry kv in dictionary)
                {
                    if (first)
                    {
                        first = !first;
                    }
                    else
                    {
                        sb.Append(", ");
                    }

                    sb.Append(kv.Key);
                    sb.Append('=');
                    sb.Append(FormatValue(kv.Value));
                }
                sb.Append('}');
                return sb.ToString();
            }

            if (value is Array array && array.Rank == 1)
            {
                var sb = new StringBuilder("[");
                var first = true;
                for (var i = 0; i < array.Length; i++)
                {
                    if (first)
                    {
                        first = !first;
                    }
                    else
                    {
                        sb.Append(", ");
                    }

                    sb.Append(FormatValue(array.GetValue(i)));
                }
                sb.Append(']');
                return sb.ToString();
            }

            return string.Format(CultureInfo.CurrentCulture, "{0}", value);
        }

        protected virtual BookDocumentRow CreateRow(BookDocument book, Row row) => new(book, this, row);
        protected virtual IDictionary<int, Column> CreateColumns() => new Dictionary<int, Column>();
        protected virtual IDictionary<int, BookDocumentRow> CreateRows() => new Dictionary<int, BookDocumentRow>();
    }
}
