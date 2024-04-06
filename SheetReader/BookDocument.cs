using System;
using System.Collections.Generic;
using System.IO;

namespace SheetReader
{
    // this class is for loading a workbook (stateful) vs enumerating it (stateless)
    public class BookDocument
    {
        private readonly List<BookDocumentSheet> _sheets = [];

        public IReadOnlyList<BookDocumentSheet> Sheets => _sheets;

        public virtual void Load(string filePath, BookFormat? format = null)
        {
            ArgumentNullException.ThrowIfNull(filePath);
            format ??= BookFormat.GetFromFileExtension(Path.GetExtension(filePath));
            ArgumentNullException.ThrowIfNull(format);

            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            format.IsStreamOwned = true;
            format.InputFilePath = filePath;
            Load(stream, format);
        }

        public virtual void Load(Stream stream, BookFormat format)
        {
            ArgumentNullException.ThrowIfNull(stream);
            _sheets.Clear();
            var book = CreateBook();
            if (book != null)
            {
                foreach (var sheet in book.EnumerateSheets(stream, format))
                {
                    var docSheet = CreateSheet(sheet);
                    if (docSheet != null)
                    {
                        _sheets.Add(docSheet);
                    }
                }
            }
        }

        protected virtual Book CreateBook() => new();
        protected virtual BookDocumentSheet CreateSheet(Sheet sheet) => new(sheet);
    }
}
