using System;

namespace SheetReader
{
    public class RowCol : IEquatable<RowCol>
    {
        public RowCol()
        {
        }

        public RowCol(int rowIndex, int columnIndex)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }

        public virtual int RowIndex { get; set; }
        public virtual int ColumnIndex { get; set; }

        public string ExcelReference => Row.GetExcelColumnName(ColumnIndex) + (RowIndex + 1).ToString();

        public override string ToString() => RowIndex + "," + ColumnIndex;
        public override bool Equals(object? obj) => Equals(obj as RowCol);
        public bool Equals(RowCol? other) => other is not null && RowIndex == other.RowIndex && ColumnIndex == other.ColumnIndex;
        public override int GetHashCode() => RowIndex.GetHashCode() ^ ColumnIndex.GetHashCode();
        public static bool operator !=(RowCol? obj1, RowCol? obj2) => !(obj1 == obj2);
        public static bool operator ==(RowCol? obj1, RowCol? obj2)
        {
            if (ReferenceEquals(obj1, obj2))
                return true;

            if (obj1 is null)
                return false;

            if (obj2 is null)
                return false;

            return obj1.Equals(obj2);
        }
    }
}
