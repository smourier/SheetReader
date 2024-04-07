namespace SheetReader.Wpf
{
    public sealed class SheetControlHitTestResult
    {
        public RowCol? RowCol { get; internal set; }
        public BookDocumentCell? Cell { get; internal set; }
        public bool IsOverRowHeader { get; internal set; }
        public bool IsOverColumnHeader { get; internal set; }
        public int? MovingColumnIndex { get; internal set; }

        public bool IsOverCorner => IsOverColumnHeader && IsOverRowHeader;
    }
}
