using System;

namespace SheetReader.Wpf
{
    public sealed class SheetSelection
    {
        private readonly SheetControl _control;
        private int _rowIndex;
        private int _rowCount = 1;
        private int _columnIndex;
        private int _columnCount = 1;
        private bool _lastRowMoveWasNegative;

        internal SheetSelection(SheetControl control)
        {
            _control = control;
        }

        public bool LastColumnMoveWasNegative { get; set; }

        public int RowIndex
        {
            get => _rowIndex;
            set
            {
                if (value < 0 || _control.Sheet == null || !_control.Sheet.LastRowIndex.HasValue)
                    return;

                if (value > _control.Sheet.LastRowIndex)
                    return;

                if (_rowIndex == value)
                    return;

                _rowIndex = value;
                var max = _control.Sheet.LastRowIndex.Value - _rowIndex + 1;
                _rowCount = Math.Max(Math.Min(_rowCount, max), 1);
                _control.OnSelectionChanged();
            }
        }

        public int ColumnIndex
        {
            get => _columnIndex;
            set
            {
                if (value < 0 || _control.Sheet == null || !_control.Sheet.LastColumnIndex.HasValue)
                    return;

                if (value > _control.Sheet.LastColumnIndex)
                    return;

                if (_columnIndex == value)
                    return;

                _columnIndex = value;
                var max = _control.Sheet.LastColumnIndex.Value - _columnIndex + 1;
                _columnCount = Math.Max(Math.Min(_columnCount, max), 1);
                _control.OnSelectionChanged();
            }
        }

        public int RowCount
        {
            get => _rowCount;
            set
            {
                if (value <= 0 || _control.Sheet == null || !_control.Sheet.LastRowIndex.HasValue)
                    return;

                if (value > _control.Sheet.LastRowIndex)
                    return;

                var max = _control.Sheet.LastRowIndex.Value - _rowIndex + 1;
                value = Math.Max(Math.Min(value, max), 1);
                if (_rowCount == value)
                    return;

                _rowCount = value;
                _control.OnSelectionChanged();
            }
        }

        public int ColumnCount
        {
            get => _columnCount;
            set
            {
                if (value < 0 || _control.Sheet == null || !_control.Sheet.LastColumnIndex.HasValue)
                    return;

                if (value == 0 || LastColumnMoveWasNegative)
                {
                    var index = ColumnIndex;
                    ColumnIndex--;
                    if (ColumnIndex == index)
                    {
                        //ColumnIndex++;
                        //value = _columnCount - 1;
                        //LastColumnMoveWasNegative = false;
                    }
                    else
                    {
                        LastColumnMoveWasNegative = true;
                        value = _columnCount + 1;
                    }
                }
                else
                {
                    LastColumnMoveWasNegative = false;
                }

                if (value > _control.Sheet.LastColumnIndex)
                    return;

                var max = _control.Sheet.LastColumnIndex.Value - _columnIndex + 1;
                value = Math.Max(Math.Min(value, max), 1);
                if (_columnCount == value)
                    return;

                _columnCount = value;
                _control.OnSelectionChanged();
            }
        }
    }
}
