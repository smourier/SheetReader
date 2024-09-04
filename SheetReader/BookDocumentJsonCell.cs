using System;
using System.Text.Json;

namespace SheetReader
{
    // minimize memory
    public class BookDocumentJsonCell : BookDocumentCell, IWithJsonElement
    {
        public BookDocumentJsonCell(Book.JsonCell cell)
            : base(cell)
        {
            ArgumentNullException.ThrowIfNull(cell);
            Value = cell.Value;
            Element = cell.Element;
        }

        public JsonElement Element { get; }
    }
}
