namespace SheetReader.Wpf
{
    public class BookDocumentStyledCell(Cell cell) : BookDocumentCell(cell)
    {
        public BookDocumentCellStyle? Style { get; set; }
    }
}
