using SheetReader.Utilities;

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

        public static BookFormat? GetFromFileExtension(string extension)
        {
            if (extension.EqualsIgnoreCase(".csv"))
                return new CsvBookFormat();

            if (extension.EqualsIgnoreCase(".xlsx"))
                return new XlsxBookFormat();

            return null;
        }
    }
}
