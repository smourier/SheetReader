using System;

namespace SheetReader
{
    [Flags]
    public enum ExportOptions
    {
        None = 0,
        StartFromFirstColumn = 1,
        StartFromFirstRow = 2,
        FirstRowIsHeader = 4,
        JsonRowsAsObject = 8,
        JsonNoDefaultCellValues = 0x10,
        JsonIndented = 0x20
    }
}
