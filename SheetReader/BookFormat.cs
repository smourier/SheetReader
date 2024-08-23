using System.Collections.Generic;
using System.IO;
using SheetReader.Utilities;

namespace SheetReader
{
    public abstract class BookFormat
    {
        private string? _name;

        protected BookFormat()
        {
        }

        public abstract BookFormatType Type { get; }
        public virtual bool IsStreamOwned { get; set; }
        public virtual string? InputFilePath { get; set; }
        public virtual string? Name
        {
            get
            {
                if (_name != null)
                    return _name;

                if (InputFilePath != null)
                {
                    try
                    {
                        return Path.GetFileNameWithoutExtension(InputFilePath);
                    }
                    catch
                    {
                        // continue;
                    }
                }

                return null;
            }
            set
            {
                if (_name == value)
                    return;

                _name = value;
            }
        }

        public static IEnumerable<string> SupportedExtensions
        {
            get
            {
                yield return ".csv";
                yield return ".xlsx";
                yield return ".json";
            }
        }

        public static BookFormat? GetFromFileExtension(string? extension)
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
