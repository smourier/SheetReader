using System;

namespace SheetReader
{
    // minimize memory
    public class BookDocumentCell
    {
        public BookDocumentCell(Cell cell)
        {
            ArgumentNullException.ThrowIfNull(cell);
            Value = cell.Value;
        }

        public virtual object? Value { get; }

        public override string ToString() => Value?.ToString() ?? string.Empty;
    }
}
