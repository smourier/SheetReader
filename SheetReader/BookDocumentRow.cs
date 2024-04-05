using System.Collections.Generic;
using System.Linq;

namespace SheetReader
{
    public sealed class BookDocumentRow
    {
        internal BookDocumentRow(Row row)
        {
            RowIndex = row.Index;
            IsHidden = !row.IsVisible;
            Cells = row.EnumerateCells().ToList().AsReadOnly();
        }

        public int RowIndex { get; }
        public bool IsHidden { get; }
        public IReadOnlyList<Cell> Cells { get; }

        public override string ToString() => RowIndex.ToString();
    }
}
