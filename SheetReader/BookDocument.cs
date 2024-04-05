using System;
using System.Collections.Generic;
using System.IO;

namespace SheetReader
{
    // this class is for loading a workbook (stateful) vs enumerating it (stateless)
    public sealed class BookDocument
    {
        private readonly List<BookDocumentSheet> _sheets = [];

        public IReadOnlyList<BookDocumentSheet> Sheets => _sheets;

        public void Load(string filePath, BookFormat? format = null)
        {
            ArgumentNullException.ThrowIfNull(filePath);
            format ??= BookFormat.GetFromFileExtension(Path.GetExtension(filePath));
            ArgumentNullException.ThrowIfNull(format);

            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            format.IsStreamOwned = true;
            format.InputFilePath = filePath;
            Load(stream, format);
        }

        public void Load(Stream stream, BookFormat format)
        {
            ArgumentNullException.ThrowIfNull(stream);
            _sheets.Clear();
            var book = new Book();
            foreach (var sheet in book.EnumerateSheets(stream, format))
            {
                var docSheet = new BookDocumentSheet(sheet);
                _sheets.Add(docSheet);
            }
        }
    }
}
