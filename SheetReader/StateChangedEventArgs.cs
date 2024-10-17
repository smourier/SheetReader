using System.ComponentModel;

namespace SheetReader
{
    public class StateChangedEventArgs(
        StateChangedType type,
        BookDocumentSheet sheet,
        BookDocumentRow? row = null,
        BookDocumentColumn? column = null,
        BookDocumentCell? cell = null) : CancelEventArgs
    {
        public StateChangedType Type { get; } = type;
        public BookDocumentSheet Sheet { get; } = sheet;
        public BookDocumentRow? Row { get; } = row;
        public BookDocumentColumn? Column { get; } = column;
        public BookDocumentCell? Cell { get; } = cell;
    }
}
