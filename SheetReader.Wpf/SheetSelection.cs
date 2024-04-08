using System;
using System.Windows;

namespace SheetReader.Wpf
{
    public class SheetSelection
    {
        public SheetSelection(SheetControl control)
        {
            ArgumentNullException.ThrowIfNull(control);
            Control = control;
        }

        public SheetSelection(SheetControl control, int rowIndex, int columnIndex)
            : this(control)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(rowIndex);
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            if (!control.Sheet.LastRowIndex.HasValue) throw new ArgumentException(null, nameof(control));
            if (!control.Sheet.LastColumnIndex.HasValue) throw new ArgumentException(null, nameof(control));
            ArgumentOutOfRangeException.ThrowIfGreaterThan(rowIndex, control.Sheet.LastRowIndex.Value);
            ArgumentOutOfRangeException.ThrowIfGreaterThan(columnIndex, control.Sheet.LastColumnIndex.Value);
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }

        public static SheetSelection? From(SheetControl control, int rowIndex, int columnIndex)
        {
            if (rowIndex < 0 || rowIndex < 0) return null;
            if (!control.Sheet.LastRowIndex.HasValue) return null;
            if (!control.Sheet.LastColumnIndex.HasValue) return null;
            if (rowIndex > control.Sheet.LastRowIndex.Value) return null;
            if (columnIndex > control.Sheet.LastColumnIndex.Value) return null;
            return new SheetSelection(control, rowIndex, columnIndex);
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
        public RowCol BottomLeft => new(BottomRight.RowIndex, TopLeft.ColumnIndex);

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

        public virtual Rect GetBounds(StyleContext? context = null)
        {
            context ??= Control.CreateStyleContext();
            if (context == null)
                throw new InvalidOperationException();

            context.RowHeight ??= Control.GetRowHeight();
            context.RowMargin ??= Control.GetRowMargin();
            context.LineSize ??= Control.GetLineSize();

            var x = context.RowFullMargin!.Value;
            for (var i = 0; i < ColumnIndex; i++)
            {
                x += Control.ColumnSettings[i].Width + context.LineSize.Value;
            }

            var width = Control.ColumnSettings[ColumnIndex].Width;
            if (ColumnExtension < 0)
            {
                for (var i = 1; i <= -ColumnExtension; i++)
                {
                    var colWidth = Control.ColumnSettings[ColumnIndex - i].Width + context.LineSize.Value;
                    x -= colWidth;
                    width += colWidth;
                }
            }
            else
            {
                for (var i = 1; i <= ColumnExtension; i++)
                {
                    width += Control.ColumnSettings[ColumnIndex + i].Width + context.LineSize.Value;
                }
            }

            var y = context.RowFullHeight!.Value + context.RowFullHeight.Value * RowIndex;
            var height = context.RowHeight.Value;
            if (RowExtension < 0)
            {
                var rowHeight = -RowExtension * context.RowFullHeight.Value;
                height += rowHeight;
                y -= rowHeight;
            }
            else
            {
                height += RowExtension * context.RowFullHeight.Value;
            }
            return new Rect(x, y, width, height);
        }

        protected internal void Update()
        {
            if (Control.Sheet == null)
                return;

            var changed = false;
            if (Control.Sheet.LastRowIndex.HasValue)
            {
                if (RowIndex > Control.Sheet.LastRowIndex.Value)
                {
                    RowIndex = Control.Sheet.LastRowIndex.Value;
                    RowExtension = 0;
                    changed = true;
                }
                else if ((RowIndex + RowExtension) > Control.Sheet.LastRowIndex.Value)
                {
                    RowExtension = Control.Sheet.LastRowIndex.Value - RowIndex;
                    changed = true;
                }
            }

            if (Control.Sheet.LastColumnIndex.HasValue)
            {
                if (ColumnIndex > Control.Sheet.LastColumnIndex.Value)
                {
                    ColumnIndex = Control.Sheet.LastColumnIndex.Value;
                    ColumnExtension = 0;
                    changed = true;
                }
                else if ((ColumnIndex + ColumnExtension) > Control.Sheet.LastColumnIndex.Value)
                {
                    ColumnExtension = Control.Sheet.LastColumnIndex.Value - ColumnIndex;
                    changed = true;
                }
            }

            if (changed)
            {
                Control.OnSelectionChanged();
            }
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
