using System;

namespace SheetReader.Wpf
{
    public class SheetSelection
    {
        public SheetSelection(SheetControl control)
        {
            ArgumentNullException.ThrowIfNull(control);
            Control = control;
        }

        public SheetControl Control { get; }
        public virtual int RowIndex { get; protected set; }
        public virtual int ColumnIndex { get; protected set; }
        public virtual int RowExtension { get; protected set; }
        public virtual int ColumnExtension { get; protected set; }

        public RowCol TopLeft
        {
            get
            {
                var rowIndex = RowIndex;
                if (RowExtension < 0)
                {
                    rowIndex += RowExtension;
                }

                var columnIndex = ColumnIndex;
                if (ColumnExtension < 0)
                {
                    columnIndex += ColumnExtension;
                }
                return new(rowIndex, columnIndex);
            }
        }

        public RowCol BottomRight
        {
            get
            {
                var rowIndex = RowIndex;
                if (RowExtension > 0)
                {
                    rowIndex += RowExtension;
                }

                var columnIndex = ColumnIndex;
                if (ColumnExtension > 0)
                {
                    columnIndex += ColumnExtension;
                }
                return new(rowIndex, columnIndex);
            }
        }

        public RowCol TopRight => new(TopLeft.RowIndex, BottomRight.ColumnIndex);
        public RowCol BottomLeft => new RowCol(BottomRight.RowIndex, TopLeft.ColumnIndex);

        public virtual void SelectRow(int rowIndex)
        {
            if (Control.Sheet == null || !Control.Sheet.LastRowIndex.HasValue)
                return;

            if (rowIndex < 0 || rowIndex >= Control.Sheet.LastRowIndex.Value)
                return;

            var changed = false;
            if (RowIndex != rowIndex)
            {
                changed = true;
                RowIndex = 0;
            }

            if (RowExtension != Control.Sheet.LastRowIndex.Value)
            {
                RowExtension = Control.Sheet.LastRowIndex.Value;
            }

            if (changed)
            {
                Control.OnSelectionChanged();
            }
        }

        public virtual void SelectColumn(int columnIndex)
        {
            if (Control.Sheet == null || !Control.Sheet.LastColumnIndex.HasValue)
                return;

            if (columnIndex < 0 || columnIndex >= Control.Sheet.LastColumnIndex.Value)
                return;

            var changed = false;
            if (ColumnIndex != columnIndex)
            {
                changed = true;
                ColumnIndex = 0;
            }

            if (ColumnExtension != Control.Sheet.LastColumnIndex.Value)
            {
                ColumnExtension = Control.Sheet.LastColumnIndex.Value;
            }

            if (changed)
            {
                Control.OnSelectionChanged();
            }
        }

        public virtual void SelectAll()
        {
            if (Control.Sheet == null || !Control.Sheet.LastRowIndex.HasValue || !Control.Sheet.LastColumnIndex.HasValue)
                return;

            var changed = false;
            if (RowIndex != 0)
            {
                changed = true;
                RowIndex = 0;
            }

            if (ColumnIndex != 0)
            {
                changed = true;
                ColumnIndex = 0;
            }

            if (RowExtension != Control.Sheet.LastRowIndex.Value)
            {
                RowExtension = Control.Sheet.LastRowIndex.Value;
            }

            if (ColumnExtension != Control.Sheet.LastColumnIndex.Value)
            {
                ColumnExtension = Control.Sheet.LastColumnIndex.Value;
            }

            if (changed)
            {
                Control.OnSelectionChanged();
            }
        }

        public void Select(RowCol rowCol) { if (rowCol == null) return; Select(rowCol.RowIndex, rowCol.ColumnIndex); }
        public virtual void Select(int rowIndex, int columnIndex)
        {
            if (Control.Sheet == null || !Control.Sheet.LastRowIndex.HasValue || !Control.Sheet.LastColumnIndex.HasValue)
                return;

            var changed = false;
            if (rowIndex < 0 || rowIndex > Control.Sheet.LastRowIndex.Value)
                return;

            if (rowIndex != RowIndex)
            {
                RowIndex = rowIndex;
                changed = true;
            }

            if (columnIndex < 0 || columnIndex > Control.Sheet.LastColumnIndex.Value)
                return;

            if (columnIndex != ColumnIndex)
            {
                ColumnIndex = columnIndex;
                changed = true;
            }

            if (RowExtension != 0)
            {
                RowExtension = 0;
                changed = true;
            }

            if (ColumnExtension != 0)
            {
                ColumnExtension = 0;
                changed = true;
            }

            if (changed)
            {
                Control.OnSelectionChanged();
            }
        }

        public void SelectTo(RowCol rowCol) { if (rowCol == null) return; SelectTo(rowCol.RowIndex, rowCol.ColumnIndex); }
        public virtual void SelectTo(int rowIndex, int columnIndex)
        {
            if (Control.Sheet == null || !Control.Sheet.LastRowIndex.HasValue || !Control.Sheet.LastColumnIndex.HasValue)
                return;

            var changed = false;
            if (rowIndex < 0 || rowIndex > Control.Sheet.LastRowIndex.Value)
                return;

            var rowExtension = rowIndex - RowIndex;
            if (rowExtension != RowExtension)
            {
                RowExtension = rowExtension;
                changed = true;
            }

            if (columnIndex < 0 || columnIndex > Control.Sheet.LastColumnIndex.Value)
                return;

            var columnExtension = columnIndex - ColumnIndex;
            if (columnExtension != ColumnExtension)
            {
                ColumnExtension = columnExtension;
                changed = true;
            }

            if (changed)
            {
                Control.OnSelectionChanged();
            }
        }

        public virtual void MoveHorizontally(int delta, bool extendSelection)
        {
            if (delta == 0 || Control.Sheet == null || !Control.Sheet.LastColumnIndex.HasValue)
                return;

            if (extendSelection)
            {
                // deal with int overflows
                var ld = (long)delta;
                var targetExtension = ColumnExtension + ld;
                var target = ColumnIndex + targetExtension;
                if (target < 0)
                {
                    ld += -target;
                }
                else if (target > Control.Sheet.LastColumnIndex.Value)
                {
                    ld -= target - Control.Sheet.LastColumnIndex.Value;
                }

                targetExtension = ColumnExtension + ld;
                if (targetExtension == ColumnExtension)
                    return;

                ColumnExtension = (int)targetExtension;
            }
            else
            {
                var target = ColumnIndex + delta;
                if (target < 0)
                {
                    delta += -target;
                }
                else if (target > Control.Sheet.LastColumnIndex.Value)
                {
                    delta -= target - Control.Sheet.LastColumnIndex.Value;
                }

                target = ColumnIndex + delta;
                ColumnIndex = target;
                ColumnExtension = 0;
                RowExtension = 0;
            }
            Control.OnSelectionChanged();
        }

        public virtual void MoveVertically(int delta, bool extendSelection)
        {
            if (delta == 0 || Control.Sheet == null || !Control.Sheet.LastRowIndex.HasValue)
                return;

            if (extendSelection)
            {
                // deal with int overflows
                var ld = (long)delta;
                var targetExtension = RowExtension + ld;
                var target = RowIndex + targetExtension;
                if (target < 0)
                {
                    ld += -target;
                }
                else if (target > Control.Sheet.LastRowIndex.Value)
                {
                    ld -= target - Control.Sheet.LastRowIndex.Value;
                }

                targetExtension = RowExtension + ld;
                if (targetExtension == RowExtension)
                    return;

                RowExtension = (int)targetExtension;
            }
            else
            {
                var target = RowIndex + delta;
                if (target < 0)
                {
                    delta += -target;
                }
                else if (target > Control.Sheet.LastRowIndex.Value)
                {
                    delta -= target - Control.Sheet.LastRowIndex.Value;
                }

                target = RowIndex + delta;
                RowIndex = target;
                RowExtension = 0;
                ColumnExtension = 0;
            }
            Control.OnSelectionChanged();
        }

        public override string ToString()
        {
            var topLeft = TopLeft;
            var bottomRight = BottomRight;
            if (bottomRight == topLeft)
                return topLeft.ExcelReference;

            return topLeft.ExcelReference + ":" + bottomRight.ExcelReference;
        }
    }
}
