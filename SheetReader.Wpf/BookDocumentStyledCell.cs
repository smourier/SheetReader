namespace SheetReader.Wpf
{
    public class BookDocumentStyledCell : BookDocumentCell
    {
        public BookDocumentStyledCell(Cell cell)
            : base(cell)
        {
        }

        public BookDocumentCellStyle? Style { get; set; }
    }
}
