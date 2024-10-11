using System;

namespace SheetReader
{
    // minimize memory
    public class BookDocumentCell : IWithValue
    {
        public BookDocumentCell(Cell cell)
        {
            ArgumentNullException.ThrowIfNull(cell);
            Value = cell.Value;
        }

        public virtual object? Value { get; set; }
        public virtual bool IsError => this is BookDocumentCellError;

        public override string ToString() => Value?.ToString() ?? string.Empty;
    }
}
