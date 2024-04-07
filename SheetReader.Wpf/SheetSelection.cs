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

        public virtual void MoveHorizontally(int delta, bool extendSelection)
        {
            if (delta == 0 || Control.Sheet == null || !Control.Sheet.LastColumnIndex.HasValue)
                return;

            var target = ColumnIndex + delta;
            if (target < 0 || target > Control.Sheet.LastColumnIndex.Value)
                return;

            if (extendSelection)
            {
                if (delta < 0)
                {
                    ColumnExtension--;
                }
                else
                {
                    ColumnExtension++;
                }
            }
            else
            {
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

            var target = RowIndex + delta;
            if (target < 0 || target > Control.Sheet.LastRowIndex.Value)
                return;

            if (extendSelection)
            {
                if (delta < 0)
                {
                    RowExtension--;
                }
                else
                {
                    RowExtension++;
                }
            }
            else
            {
                RowIndex = target;
                RowExtension = 0;
                ColumnExtension = 0;
            }
            Control.OnSelectionChanged();
        }
    }
}
