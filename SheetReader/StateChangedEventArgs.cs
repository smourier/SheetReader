using System.ComponentModel;

namespace SheetReader
{
    public class StateChangedEventArgs(
        StateChangedType type,
        BookDocumentSheet sheet,
        BookDocumentRow? row = null,
        Column? column = null,
        BookDocumentCell? cell = null) : CancelEventArgs
    {
        public StateChangedType Type { get; } = type;
        public BookDocumentSheet Sheet { get; } = sheet;
        public BookDocumentRow? Row { get; } = row;
        public Column? Column { get; } = column;
        public BookDocumentCell? Cell { get; } = cell;
    }
}
