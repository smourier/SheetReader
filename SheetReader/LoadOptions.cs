using System;

namespace SheetReader
{
    [Flags]
    public enum LoadOptions
    {
        None = 0x0,
        FirstRowDefinesColumns = 0x1, // sometimes, this can be guessed
    }
}
