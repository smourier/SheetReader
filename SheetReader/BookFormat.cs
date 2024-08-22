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
            if (Extensions.EqualsIgnoreCase(extension, ".csv"))
                return new CsvBookFormat();

            if (Extensions.EqualsIgnoreCase(extension, ".xlsx"))
                return new XlsxBookFormat();

            if (Extensions.EqualsIgnoreCase(extension, ".json"))
            {
                return new JsonBookFormat();
            }
            return null;
        }
    }
}
