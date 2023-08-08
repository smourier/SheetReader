namespace SheetReader
{
    public abstract class BookFormat
    {
        protected BookFormat()
        {
        }

        public abstract BookFormatType Type { get; }
        public virtual bool IsStreamOwned { get; set; }
        public virtual string? InputFilePath { get; set; }
    }
}
