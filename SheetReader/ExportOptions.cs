﻿using System;

namespace SheetReader
{
    [Flags]
    public enum ExportOptions
    {
        None = 0x0,
        StartFromFirstColumn = 0x1,
        StartFromFirstRow = 0x2,
        FirstRowIsHeader = 0x4,

        // json only
        JsonRowsAsObject = 0x8,
        JsonNoDefaultCellValues = 0x10,
        JsonIndented = 0x20,

        // csv only
        CsvWriteColumns = 0x40,
    }
}