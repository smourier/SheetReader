﻿using System;

namespace SheetReader
{
    [Flags]
    public enum JsonBookOptions
    {
        None = 0x0,
        ParseDates = 0x1,
        DontAddColumnsFromRows = 0x2,

        Default = ParseDates,
    }
}
