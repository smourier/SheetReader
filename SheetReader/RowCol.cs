namespace SheetReader
{
    public class RowCol
    {
        public virtual int RowIndex { get; set; }
        public virtual int ColumnIndex { get; set; }

        public override string ToString() => RowIndex + "," + ColumnIndex;
    }
}
