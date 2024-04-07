namespace SheetReader
{
    public class RowCol
    {
        public virtual int RowIndex { get; set; }
        public virtual int ColumnIndex { get; set; }

        public string ExcelReference => Row.GetExcelColumnName(ColumnIndex) + (RowIndex + 1).ToString();
        public override string ToString() => RowIndex + "," + ColumnIndex;
    }
}
