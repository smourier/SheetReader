/*
MIT License

Copyright (c) 2023-2024 Simon Mourier

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/
/*
AssemblyVersion: 2.0.0.0
AssemblyFileVersion: 2.0.0.0
*/
using Extensions = SheetReader.Utilities.Extensions;
using SheetReader.Utilities;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Compression;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Xml;
using System;

namespace SheetReader
{
    public class Book
    {
        protected static XmlNamespaceManager XmlNamespaceManager { get; } = BuildMgr();
        private const string _relDoc = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string _rels = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string _r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string _ws = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        private const string _main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string _ss = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
        private const string _styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

        private const string _oxRelDoc = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
        private const string _oxStyles = "http://purl.oclc.org/ooxml/officeDocument/relationships/styles";
        private const string _oxSs = "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings";
        private const string _oxWs = "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet";
        private const string _oxMain = "http://purl.oclc.org/ooxml/spreadsheetml/main";
        private const string _oxR = "http://purl.oclc.org/ooxml/officeDocument/relationships";

        private static XmlNamespaceManager BuildMgr()
        {
            var mgr = new XmlNamespaceManager(new NameTable());
            mgr.AddNamespace("rdoc", _relDoc);
            mgr.AddNamespace("rels", _rels);
            mgr.AddNamespace("r", _r);
            mgr.AddNamespace("ws", _ws);
            mgr.AddNamespace("ss", _ss);
            mgr.AddNamespace("styles", _styles);
            mgr.AddNamespace("main", _main);
            mgr.AddNamespace("oxmain", _oxMain);
            return mgr;
        }

        public static bool IsSupportedFileExtension(string extension) => BookFormat.GetFromFileExtension(extension) != null;

        public virtual IEnumerable<Sheet> EnumerateSheets(string filePath, BookFormat? format = null)
        {
            ArgumentNullException.ThrowIfNull(filePath);
            format ??= BookFormat.GetFromFileExtension(Path.GetExtension(filePath));
            ArgumentNullException.ThrowIfNull(format);

            var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            format.IsStreamOwned = true;
            format.InputFilePath = filePath;
            return EnumerateSheets(stream, format);
        }

        public virtual IEnumerable<Sheet> EnumerateSheets(Stream stream, BookFormat format)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(format);

            if (format is CsvBookFormat csv)
                return EnumerateCsvSheets(stream, csv);

            if (format is XlsxBookFormat xlsx)
                return EnumerateXlsxSheets(stream, xlsx);

            throw new NotSupportedException();
        }

        protected virtual IEnumerable<Sheet> EnumerateCsvSheets(Stream stream, CsvBookFormat format)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(format);
            var sheet = CreateCsvSheet(stream, format);
            if (sheet != null)
            {
                yield return sheet;
            }

            if (format.IsStreamOwned)
            {
                stream.Dispose();
            }
        }

        protected virtual CsvSheet CreateCsvSheet(Stream stream, CsvBookFormat format) => new(stream, format);

        protected virtual IEnumerable<Sheet> EnumerateXlsxSheets(Stream stream, XlsxBookFormat format)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(format);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, !format.IsStreamOwned);
            var relsEntry = archive.GetEntry("_rels/.rels");
            if (relsEntry == null)
                yield break;

            using var relsStream = relsEntry.Open();
            var relsDoc = XDocument.Load(relsStream);
            var workbookEntryName = XlsxBook.GetTarget(relsDoc, null, _relDoc, _oxRelDoc);
            if (workbookEntryName == null)
                yield break;

            var workbookEntry = archive.GetEntry(workbookEntryName);
            if (workbookEntry == null)
                yield break;

            var relPath = Path.GetDirectoryName(workbookEntryName);
            if (relPath == null)
                yield break;

            var workbookRelsEntryName = Extensions.NormPath(Path.Combine(relPath, "_rels", Path.GetFileName(workbookEntryName) + ".rels"));
            var workbookRelsEntry = archive.GetEntry(workbookRelsEntryName);
            if (workbookRelsEntry == null)
                yield break;

            using var workbookStream = workbookEntry.Open();
            var workBookDoc = XDocument.Load(workbookStream);
            using var workbookRelsStream = workbookRelsEntry.Open();
            var workBookRelsDoc = XDocument.Load(workbookRelsStream);

            var wb = CreateXlsxBook(archive, relPath, workBookDoc, workBookRelsDoc);
            if (wb != null)
            {
                foreach (var sheet in wb.EnumerateSheets())
                {
                    yield return sheet;
                }
            }
        }

        protected virtual XlsxBook CreateXlsxBook(ZipArchive archive, string relativePath, XDocument workbookDocument, XDocument relsDocument) => new(archive, relativePath, workbookDocument, relsDocument);

        protected class XlsxBook
        {
            public XlsxBook(ZipArchive archive, string relativePath, XDocument workbookDocument, XDocument relsDocument)
            {
                ArgumentNullException.ThrowIfNull(archive);
                ArgumentNullException.ThrowIfNull(relativePath);
                ArgumentNullException.ThrowIfNull(workbookDocument);
                ArgumentNullException.ThrowIfNull(relsDocument);
                Archive = archive;
                RelativePath = relativePath;
                WorkbookDocument = workbookDocument;
                RelsDocument = relsDocument;

                var sharedStrings = new List<string>();
                var sharedStringsEntryName = GetTarget(relsDocument, null, _ss, _oxSs);
                if (sharedStringsEntryName != null)
                {
                    var sharedStringsEntry = archive.GetEntry(Extensions.NormPath(Path.Combine(relativePath, sharedStringsEntryName)));
                    if (sharedStringsEntry != null)
                    {
                        using var sharedStringsStream = sharedStringsEntry.Open();
                        var sharedStringsDoc = XDocument.Load(sharedStringsStream);
                        string xpath;
                        if (IsOpenXml(sharedStringsDoc.Root))
                        {
                            xpath = "oxmain:sst/oxmain:si";
                        }
                        else
                        {
                            xpath = "main:sst/main:si";
                        }

                        foreach (var stringElement in sharedStringsDoc.XPathSelectElements(xpath, XmlNamespaceManager))
                        {
                            // value will get all text beneath, including rtf
                            sharedStrings.Add(stringElement.Value);
                        }
                    }
                }
                SharedStrings = sharedStrings.AsReadOnly();

                var formats = new Dictionary<int, XlsxFormat>();
                var stylesEntryName = GetTarget(relsDocument, null, _styles, _oxStyles);
                if (stylesEntryName != null)
                {
                    var stylesEntry = archive.GetEntry(Extensions.NormPath(Path.Combine(relativePath, stylesEntryName)));
                    if (stylesEntry != null)
                    {
                        using var stylesStream = stylesEntry.Open();
                        var stylesDoc = XDocument.Load(stylesStream);
                        string xpath;
                        if (IsOpenXml(stylesDoc.Root))
                        {
                            xpath = "oxmain:styleSheet/oxmain:cellXfs/oxmain:xf";
                        }
                        else
                        {
                            xpath = "main:styleSheet/main:cellXfs/main:xf";
                        }

                        foreach (var formatElement in stylesDoc.XPathSelectElements(xpath, XmlNamespaceManager))
                        {
                            var xformat = CreateXlsxFormat(formatElement);
                            if (xformat != null)
                            {
                                formats[formats.Count] = xformat;
                            }
                        }
                    }
                }
                Formats = formats.AsReadOnly();
            }

            public ZipArchive Archive { get; }
            public string RelativePath { get; }
            public XDocument WorkbookDocument { get; }
            public XDocument RelsDocument { get; }
            public IReadOnlyList<string> SharedStrings { get; }
            public IReadOnlyDictionary<int, XlsxFormat> Formats { get; }

            protected virtual XlsxFormat CreateXlsxFormat(XElement element) => new(this, element);
            protected virtual XlsxSheet CreateXlsxSheet(XElement element, XmlReader reader) => new(this, element, reader);

            internal static bool IsMainNamespace(string ns) => ns == _main || ns == _oxMain;
            internal static bool IsOpenXml(XElement? element) => element?.Name.Namespace.NamespaceName == _oxMain;
            internal static string? GetTarget(XNode node, string? id, params string[] types)
            {
                if (id == null)
                {
                    foreach (var type in types)
                    {
                        var target = node.XPathEvaluate($"string(rels:Relationships/rels:Relationship[@Type='{type}']/@Target)", XmlNamespaceManager) as string;
                        if (!string.IsNullOrEmpty(target))
                            return target;
                    }
                }

                foreach (var type in types)
                {
                    var target = node.XPathEvaluate($"string(rels:Relationships/rels:Relationship[@Type='{type}' and @Id='{id}']/@Target)", XmlNamespaceManager) as string;
                    if (!string.IsNullOrEmpty(target))
                        return target;
                }
                return null;
            }

            public virtual IEnumerable<XlsxSheet> EnumerateSheets()
            {
                string xpath;
                var openXml = IsOpenXml(WorkbookDocument.Root);
                if (openXml)
                {
                    xpath = "oxmain:workbook/oxmain:sheets/oxmain:sheet";
                }
                else
                {
                    xpath = "main:workbook/main:sheets/main:sheet";
                }

                foreach (var sheetElement in WorkbookDocument.XPathSelectElements(xpath, XmlNamespaceManager))
                {
                    if (sheetElement == null)
                        continue;

                    var id = sheetElement.Attribute(XName.Get("id", openXml ? _oxR : _r))?.Value;
                    if (id == null)
                        continue;

                    var sheetEntryName = GetTarget(RelsDocument, id, _ws, _oxWs);
                    if (sheetEntryName == null)
                        continue;

                    sheetEntryName = Extensions.NormPath(Path.Combine(RelativePath, sheetEntryName));
                    var sheetEntry = Archive.GetEntry(sheetEntryName);
                    if (sheetEntry == null)
                        continue;

                    using var sheetStream = sheetEntry.Open();
                    using var reader = XmlReader.Create(sheetStream);
                    var sheet = CreateXlsxSheet(sheetElement, reader);
                    if (sheet != null)
                        yield return sheet;
                }
            }
        }

        protected class XlsxFormat
        {
            public XlsxFormat(XlsxBook book, XElement element)
            {
                ArgumentNullException.ThrowIfNull(book);
                ArgumentNullException.ThrowIfNull(element);
                Book = book;
                var id = element.Attribute("numFmtId")?.Value;
                if (id != null && int.TryParse(id, out var i) && i >= 0)
                {
                    FormatId = i;
                }
            }

            public XlsxBook Book { get; }
            public int FormatId { get; }

            public virtual bool TryParseValue(string rawWalue, out object? value)
            {
                value = null;
                if (rawWalue == null)
                    return false;

                // https://github.com/closedxml/closedxml/wiki/NumberFormatId-Lookup-Table
                switch (FormatId)
                {
                    case 14:
                    case 15:
                    case 16:
                    case 17:
                    case 22:
                        if (double.TryParse(rawWalue, CultureInfo.InvariantCulture, out var dbl))
                        {
                            try
                            {
                                value = DateTime.FromOADate(dbl);
                                return true;
                            }
                            catch
                            {
                                // continue;
                            }
                        }
                        break;
                }

                return false;
            }
        }

        protected class XlsxSheet : Sheet
        {
            public XlsxSheet(XlsxBook book, XElement element, XmlReader reader)
            {
                ArgumentNullException.ThrowIfNull(book);
                ArgumentNullException.ThrowIfNull(element);
                ArgumentNullException.ThrowIfNull(reader);
                Book = book;
                Element = element;
                Reader = reader;
                Name = element.Attribute("name")?.Value!;
                var state = element.Attribute("state")?.Value;
                if (state.EqualsIgnoreCase("hidden"))
                {
                    IsVisible = false;
                }
            }

            public XlsxBook Book { get; }
            public XElement Element { get; }
            public XmlReader Reader { get; }
            public IDictionary<int, Column> Columns { get; } = new Dictionary<int, Column>();

            public override IEnumerable<Column> EnumerateColumns() => Columns.Values.OrderBy(c => c.Index);

            public override IEnumerable<Row> EnumerateRows()
            {
                while (Reader.Read())
                {
                    if (Reader.NodeType != XmlNodeType.Element)
                        continue;

                    if (Reader.LocalName == "sheetData" && XlsxBook.IsMainNamespace(Reader.NamespaceURI))
                    {
                        while (Reader.Read())
                        {
                            if (Reader.NodeType == XmlNodeType.EndElement)
                            {
                                if (Reader.LocalName == "sheetData" && XlsxBook.IsMainNamespace(Reader.NamespaceURI))
                                    yield break;
                            }

                            if (Reader.LocalName == "row" && XlsxBook.IsMainNamespace(Reader.NamespaceURI))
                            {
                                var r = Reader.GetAttribute("r");
                                if (r == null || !int.TryParse(r, out var rowIndex) || rowIndex <= 0)
                                    continue;

                                var row = CreateRow();
                                if (row != null)
                                {
                                    row.Index = rowIndex - 1;
                                    var hidden = Reader.GetAttribute("hidden");
                                    if (hidden != null && hidden.EqualsIgnoreCase("true") || hidden == "1")
                                    {
                                        row.IsVisible = false;
                                    }
                                    yield return row;
                                }
                            }
                        }
                    }
                }
            }

            protected virtual XlsxRow CreateRow() => new(this);
            protected virtual internal XlsxCell CreateCell(XlsxRow row) => new(row);
        }

        protected class XlsxRow : Row
        {
            public XlsxRow(XlsxSheet sheet)
            {
                ArgumentNullException.ThrowIfNull(sheet);
                Sheet = sheet;
            }

            public XlsxSheet Sheet { get; }

            public override IEnumerable<Cell> EnumerateCells()
            {
                while (Sheet.Reader.Read())
                {
                    if (Sheet.Reader.NodeType == XmlNodeType.EndElement)
                    {
                        if (Sheet.Reader.LocalName == "row" && XlsxBook.IsMainNamespace(Sheet.Reader.NamespaceURI))
                            yield break;
                    }

                    if (Sheet.Reader.LocalName == "c" && XlsxBook.IsMainNamespace(Sheet.Reader.NamespaceURI))
                    {
                        var cell = Sheet.CreateCell(this);
                        if (cell != null)
                            yield return cell;
                    }
                }
            }
        }

        protected class XlsxCell : Cell
        {
            public XlsxCell(XlsxRow row)
            {
                ArgumentNullException.ThrowIfNull(row);
                Row = row;
                var reader = Row.Sheet.Reader;
                var type = reader.GetAttribute("t");
                var format = reader.GetAttribute("s");
                var reference = reader.GetAttribute("r");
                if (reference != null)
                {
                    var index = 0;
                    foreach (var c in reference)
                    {
                        if (c < 'A' || c > 'Z')
                            break;

                        index = 26 * index + c - 'A' + 1;
                    }
                    ColumnIndex = index - 1;
                }

                if (row.Sheet.Columns.Count <= ColumnIndex)
                {
                    if (!row.Sheet.Columns.ContainsKey(ColumnIndex))
                    {
                        var column = row.Sheet.CreateColumn();
                        column.Index = ColumnIndex;
                        row.Sheet.Columns.Add(ColumnIndex, column);
                    }
                }

                if (!reader.IsEmptyElement)
                {
                    while (reader.Read())
                    {
                        if (reader.NodeType == XmlNodeType.EndElement)
                        {
                            if (reader.LocalName == "c" && XlsxBook.IsMainNamespace(reader.NamespaceURI))
                                break;
                        }

                        if (reader.LocalName == "v" && XlsxBook.IsMainNamespace(reader.NamespaceURI))
                        {
                            RawValue = reader.ReadElementContentAsString();
                            break;
                        }
                    }
                }

                switch (type)
                {
                    case "e":
                        IsError = true;
                        break;

                    case "s":
                        if (int.TryParse(RawValue, out var i) && i >= 0 && i < Row.Sheet.Book.SharedStrings.Count)
                        {
                            Value = Row.Sheet.Book.SharedStrings[i];
                        }
                        break;

                    case "d":
                        if (RawValue != null && DateTime.TryParseExact(RawValue, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite, out var dt))
                        {
                            Value = dt;
                        }
                        break;

                    case "b":
                        Value = RawValue == "1";
                        break;
                }

                if (Value == null && RawValue != null)
                {
                    if (format != null &&
                        int.TryParse(format, out var fmtId) &&
                        Row.Sheet.Book.Formats.TryGetValue(fmtId, out var xformat) &&
                        xformat.TryParseValue(RawValue, out var formatted))
                    {
                        Value = formatted;
                    }
                    else
                    {
                        Value = RawValue;
                    }
                }
            }

            public XlsxRow Row { get; }
        }

        protected class CsvRow : Row
        {
            public CsvRow(IEnumerator<string> cells)
            {
                ArgumentNullException.ThrowIfNull(cells);
                Cells = cells;
            }

            public IEnumerator<string> Cells { get; }

            public override IEnumerable<Cell> EnumerateCells()
            {
                var index = 0;
                while (Cells.MoveNext())
                {
                    var cell = CreateCell();
                    cell.ColumnIndex = index;
                    cell.Value = Cells.Current;
                    cell.RawValue = Cells.Current;
                    if (cell != null)
                    {
                        yield return cell;
                        index++;
                    }
                }
            }

            protected virtual Cell CreateCell() => new();
        }

        protected class CsvSheet : Sheet, IDisposable
        {
            private bool _disposedValue;

            public CsvSheet(Stream stream, CsvBookFormat format)
            {
                ArgumentNullException.ThrowIfNull(stream);
                ArgumentNullException.ThrowIfNull(format);
                if (format.InputFilePath != null)
                {
                    Name = Path.GetFileNameWithoutExtension(format.InputFilePath);
                }
                Reader = new CsvReader(stream, format.AllowCharacterAmbiguity, format.ReadHeaderRow, format.Quote, format.Separator, format.Encoding);
            }

            public CsvReader Reader { get; }

            public override IEnumerable<Row> EnumerateRows()
            {
                var rowIndex = 0;
                while (!Reader.EndOfStream)
                {
                    var cells = Reader.ReadRow().GetEnumerator();
                    var row = CreateRow(cells);
                    if (row != null)
                    {
                        row.Index = rowIndex;
                        yield return row;
                        rowIndex++;
                    }
                }
            }

            public override IEnumerable<Column> EnumerateColumns()
            {
                for (var i = 0; i < Reader.Columns.Count; i++)
                {
                    var column = CreateColumn();
                    if (column != null)
                    {
                        column.Name = Reader.Columns[i];
                        column.Index = i;
                        yield return column;
                    }
                }
            }

            protected virtual CsvRow CreateRow(IEnumerator<string> cells) => new(cells);

            protected virtual void Dispose(bool disposing)
            {
                if (!_disposedValue)
                {
                    if (disposing)
                    {
                        // dispose managed state (managed objects)
                        Reader.Dispose();
                    }

                    // free unmanaged resources (unmanaged objects) and override finalizer
                    // set large fields to null
                    _disposedValue = true;
                }
            }

            // // override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
            // ~CsvSheet()
            // {
            //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            //     Dispose(disposing: false);
            // }

            public void Dispose()
            {
                // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
                Dispose(disposing: true);
                GC.SuppressFinalize(this);
            }
        }
    }
}

namespace SheetReader
{
    // this class is for loading a workbook (stateful) vs enumerating it (stateless)
    public class BookDocument
    {
        private readonly IList<BookDocumentSheet> _sheets;

        public BookDocument()
        {
            _sheets = CreateSheets();
            if (_sheets == null)
                throw new InvalidOperationException();
        }

        public IList<BookDocumentSheet> Sheets => _sheets;

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
        protected virtual IList<BookDocumentSheet> CreateSheets() => [];
    }
}

namespace SheetReader
{
    // minimize memory
    public class BookDocumentCell
    {
        public BookDocumentCell(Cell cell)
        {
            ArgumentNullException.ThrowIfNull(cell);
            Value = cell.Value;
        }

        public virtual object? Value { get; }

        public override string ToString() => Value?.ToString() ?? string.Empty;
    }
}

namespace SheetReader
{
    public class BookDocumentCellError(Cell cell) : BookDocumentCell(cell)
    {
    }
}

namespace SheetReader
{
    public class BookDocumentRow
    {
        private readonly IDictionary<int, BookDocumentCell> _cells;

        public BookDocumentRow(Row row)
        {
            ArgumentNullException.ThrowIfNull(row);
            _cells = CreateCells();
            if (_cells == null)
                throw new InvalidOperationException();

            RowIndex = row.Index;
            IsHidden = !row.IsVisible;
            foreach (var cell in row.EnumerateCells())
            {
                var bdCell = CreateCell(cell);
                if (bdCell == null)
                    continue;

                _cells[cell.ColumnIndex] = CreateCell(cell);

                if (LastCellIndex == null || row.Index > LastCellIndex)
                {
                    LastCellIndex = row.Index;
                }

                if (FirstCellIndex == null || row.Index < FirstCellIndex)
                {
                    FirstCellIndex = row.Index;
                }
            }
        }

        public int RowIndex { get; }
        public bool IsHidden { get; }
        public int? FirstCellIndex { get; }
        public int? LastCellIndex { get; }
        public IDictionary<int, BookDocumentCell> Cells => _cells;

        protected virtual IDictionary<int, BookDocumentCell> CreateCells() => new Dictionary<int, BookDocumentCell>();
        protected virtual BookDocumentCell CreateCell(Cell cell)
        {
            ArgumentNullException.ThrowIfNull(cell);
            return cell.IsError ? new BookDocumentCellError(cell) : new BookDocumentCell(cell);
        }

        public override string ToString() => RowIndex.ToString();
    }
}

namespace SheetReader
{
    public class BookDocumentSheet
    {
        private readonly IDictionary<int, BookDocumentRow> _rows;
        private readonly IDictionary<int, Column> _columns;

        public BookDocumentSheet(Sheet sheet)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            _rows = CreateRows();
            _columns = CreateColumns();
            if (_rows == null || _columns == null)
                throw new InvalidOperationException();

            Name = sheet.Name ?? string.Empty;
            IsHidden = !sheet.IsVisible;

            foreach (var row in sheet.EnumerateRows())
            {
                var rowData = CreateRow(row);
                _rows[row.Index] = rowData;

                if (LastRowIndex == null || row.Index > LastRowIndex)
                {
                    LastRowIndex = row.Index;
                }

                if (FirstRowIndex == null || row.Index < FirstRowIndex)
                {
                    FirstRowIndex = row.Index;
                }
            }

            foreach (var col in sheet.EnumerateColumns())
            {
                _columns[col.Index] = col;
                if (LastColumnIndex == null || col.Index > LastColumnIndex)
                {
                    LastColumnIndex = col.Index;
                }

                if (FirstColumnIndex == null || col.Index < FirstColumnIndex)
                {
                    FirstColumnIndex = col.Index;
                }
            }
        }

        public string Name { get; }
        public bool IsHidden { get; }
        public int? FirstColumnIndex { get; }
        public int? LastColumnIndex { get; }
        public int? FirstRowIndex { get; }
        public int? LastRowIndex { get; }
        public IDictionary<int, BookDocumentRow> Rows => _rows;
        public IDictionary<int, Column> Columns => _columns;

        public override string ToString() => Name;

        public BookDocumentCell? GetCell(RowCol? rowCol)
        {
            if (rowCol == null)
                return null;

            return GetCell(rowCol.RowIndex, rowCol.ColumnIndex);
        }

        public virtual BookDocumentCell? GetCell(int rowIndex, int columnIndex)
        {
            if (!_rows.TryGetValue(rowIndex, out var row))
                return null;

            row.Cells.TryGetValue(columnIndex, out var cell);
            return cell;
        }

        protected virtual BookDocumentRow CreateRow(Row row) => new(row);
        protected virtual IDictionary<int, Column> CreateColumns() => new Dictionary<int, Column>();
        protected virtual IDictionary<int, BookDocumentRow> CreateRows() => new Dictionary<int, BookDocumentRow>();
    }
}

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

namespace SheetReader
{
    public enum BookFormatType
    {
        Automatic,
        Csv,
        Xlsx,
    }
}

namespace SheetReader
{
    public class Cell
    {
        public virtual int ColumnIndex { get; set; }
        public virtual object? Value { get; set; }
        public virtual bool IsError { get; set; }
        public virtual string? RawValue { get; set; }

        public override string ToString() => Value?.ToString() ?? string.Empty;
    }
}

namespace SheetReader
{
    public class Column
    {
        public virtual int Index { get; set; }
        public virtual string? Name { get; set; }

        public override string ToString() => Name.Nullify() ?? Index.ToString();
    }
}

namespace SheetReader
{
    public class CsvBookFormat : BookFormat
    {
        public override BookFormatType Type => BookFormatType.Csv;

        public virtual bool AllowCharacterAmbiguity { get; set; } = false;
        public virtual bool ReadHeaderRow { get; set; } = true;
        public virtual char Quote { get; set; } = '"';
        public virtual char Separator { get; set; } = ';';
        public virtual Encoding? Encoding { get; set; }
    }
}

namespace SheetReader
{
    public class CsvReader : IDisposable
    {
        private readonly List<string> _columns = new();
        private Stream? _ownedStream;
        private bool _disposedValue;

        public CsvReader(string filePath, bool allowCharacterAmbiguity, bool readHeaderRow, char quote, char separator, Encoding? encoding)
        {
            ArgumentNullException.ThrowIfNull(filePath);

            var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            _ownedStream = stream;
            Initialize(stream, allowCharacterAmbiguity, readHeaderRow, quote, separator, encoding);
        }

        public CsvReader(Stream stream, bool allowCharacterAmbiguity, bool readHeaderRow, char quote, char separator, Encoding? encoding)
        {
            Initialize(stream, allowCharacterAmbiguity, readHeaderRow, quote, separator, encoding);
        }

        public CsvReader(TextReader reader, bool allowCharacterAmbiguity, bool readHeaderRow, char quote, char separator)
        {
            ArgumentNullException.ThrowIfNull(reader);
            BaseReader = reader;
            InitializeReader(allowCharacterAmbiguity, readHeaderRow, quote, separator);
        }

        public TextReader? BaseReader { get; private set; }
        public IReadOnlyList<string> Columns => _columns;
        public virtual char Separator { get; protected set; }
        public virtual char Quote { get; protected set; }
        public virtual Encoding? Encoding { get; protected set; }
        public virtual int LineNumber { get; protected set; }
        public virtual int ColumnNumber { get; protected set; }
        public virtual int RowNumber { get; protected set; }
        public virtual bool ReadHeaderRow { get; protected set; }
        protected bool IsReaderOwned { get; set; }

        public bool EndOfStream
        {
            get
            {
                if (BaseReader == null)
                    return true;

                if (BaseReader is StreamReader sr)
                    return sr.EndOfStream;

                return BaseReader.Peek() < 0;
            }
        }

        public static IEnumerable<string> ReadLine(string line) => ReadLine(line, '"', ';');
        public static IEnumerable<string> ReadLine(string line, char quote, char separator)
        {
            if (line == null)
                yield break;

            using var reader = new StringReader(line);
            using var csvReader = new CsvReader(reader, true, false, quote, separator);
            foreach (string cell in csvReader.ReadRow())
            {
                yield return cell;
            }
        }

        [MemberNotNull(nameof(BaseReader), nameof(Encoding))]
        protected void Initialize(Stream stream, bool allowCharacterAmbiguity, bool readHeaderRow, char quote, char separator, Encoding? encoding)
        {
            ArgumentNullException.ThrowIfNull(stream);
            IsReaderOwned = true;
            Encoding = encoding ?? Encoding.UTF8;
            BaseReader = new StreamReader(stream, Encoding, true);
            InitializeReader(allowCharacterAmbiguity, readHeaderRow, quote, separator);
        }

        protected void InitializeReader(bool allowCharacterAmbiguity, bool readHeaderRow, char quote, char separator)
        {
            Quote = quote;
            Separator = separator;
            BaseReader!.Peek(); // force encoding detection
            DetermineFormat(allowCharacterAmbiguity);
            ReadHeaderRow = readHeaderRow;
            if (readHeaderRow)
            {
                ReadHeader();
            }
        }

        public virtual IEnumerable<string> ReadRow()
        {
            if (BaseReader == null)
                yield break;

            var inQuote = false;
            int next;
            char c;
            var cell = new StringBuilder();
            var hasCell = true;
            ColumnNumber = 0;
            do
            {
                var i = BaseReader.Read();
                if (i < 0)
                {
                    if (hasCell)
                        yield return cell.ToString();

                    break;
                }

                ColumnNumber++;
                next = BaseReader.Peek();

                c = (char)i;
                if (inQuote)
                {
                    if (c == '\n')
                    {
                        LineNumber++;
                    }
                    else if (c == '\r')
                    {
                        if (next == '\n')
                        {
                            hasCell = true;
                            cell.Append(c);
                            ColumnNumber++;
                            c = (char)BaseReader.Read();
                        }
                        LineNumber++;
                    }

                    if (c == Quote)
                    {
                        if (next == Quote)
                        {
                            BaseReader.Read();
                            hasCell = true;
                            cell.Append(Quote);
                        }
                        else
                        {
                            inQuote = false;
                        }
                        continue;
                    }

                    hasCell = true;
                    cell.Append(c);
                }
                else
                {
                    if (c == '\n')
                    {
                        if (hasCell)
                        {
                            yield return cell.ToString();
                            cell.Length = 0;
                        }
                        LineNumber++;
                        break;
                    }

                    if (c == '\r')
                    {
                        if (next == '\n')
                        {
                            BaseReader.Read();
                            ColumnNumber++;
                        }

                        if (hasCell)
                        {
                            yield return cell.ToString();
                            cell.Length = 0;
                        }
                        LineNumber++;
                        break;
                    }

                    if (c == Quote)
                    {
                        if (next == Quote)
                        {
                            BaseReader.Read();
                            hasCell = true;
                            cell.Append(Quote);
                        }
                        else
                        {
                            inQuote = true;
                        }
                        continue;
                    }

                    if (c == Separator)
                    {
                        if (hasCell)
                        {
                            yield return cell.ToString();
                            cell.Length = 0;
                        }
                        // note: we keep hasCell to true
                        continue;
                    }

                    hasCell = true;
                    cell.Append(c);
                }
            }
            while (true);
            RowNumber++;
        }

        protected virtual void ReadHeader()
        {
            if (BaseReader == null)
                throw new InvalidOperationException();

            foreach (var column in ReadRow())
            {
                _columns.Add(column);
            }
        }

        protected virtual void DetermineFormat(bool allowCharacterAmbiguity)
        {
            if (BaseReader == null)
                return;

            var sr = BaseReader as StreamReader;
            if (sr != null && sr.CurrentEncoding is UTF8Encoding)
                return;

            if (sr != null && sr.CurrentEncoding is UnicodeEncoding)
            {
                Separator = '\t';
                return;
            }

            if (!allowCharacterAmbiguity && sr != null)
                throw new SheetReaderException($"0002: CSV has ambiguous file encoding: '{sr.CurrentEncoding.WebName}'.");
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    // dispose managed state (managed objects)
                    var br = BaseReader;
                    BaseReader = null;
                    if (IsReaderOwned)
                    {
                        br?.Dispose();
                    }

                    _ownedStream?.Dispose();
                    _ownedStream = null;
                }

                _disposedValue = true;
            }
        }

        public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
    }
}

namespace SheetReader
{
    public abstract class Row
    {
        public virtual int Index { get; set; }
        public virtual bool IsVisible { get; set; } = true;

        public abstract IEnumerable<Cell> EnumerateCells();

        public static string GetExcelColumnName(int index)
        {
            index++;
            var name = string.Empty;
            while (index > 0)
            {
                var mod = (index - 1) % 26;
                name = (char)('A' + mod) + name;
                index = (index - mod) / 26;
            }
            return name;
        }

        public override string ToString() => Index.ToString();
    }
}

namespace SheetReader
{
    public class RowCol
    {
        public virtual int RowIndex { get; set; }
        public virtual int ColumnIndex { get; set; }

        public string ExcelReference => Row.GetExcelColumnName(ColumnIndex) + (RowIndex + 1).ToString();
        public override string ToString() => RowIndex + "," + ColumnIndex;
    }
}

namespace SheetReader
{
    public abstract class Sheet
    {
        public virtual string? Name { get; set; }
        public virtual bool IsVisible { get; set; } = true;

        public abstract IEnumerable<Column> EnumerateColumns();
        public abstract IEnumerable<Row> EnumerateRows();

        protected internal virtual Column CreateColumn() => new();

        public override string ToString() => Name ?? string.Empty;
    }
}

namespace SheetReader
{
    [Serializable]
    public class SheetReaderException : Exception
    {
        public const string Prefix = "SHR";

        public SheetReaderException()
            : base(Prefix + "0001: SheetReader exception.")
        {
        }

        public SheetReaderException(string message)
            : base(Prefix + ":" + message)
        {
        }

        public SheetReaderException(Exception innerException)
            : base(null, innerException)
        {
        }

        public SheetReaderException(string message, Exception innerException)
            : base(Prefix + ":" + message, innerException)
        {
        }

        public int Code => GetCode(Message);

        public static int GetCode(string message)
        {
            if (message == null)
                return -1;

            if (!message.StartsWith(Prefix, StringComparison.Ordinal))
                return -1;

            var pos = message.IndexOf(':', Prefix.Length);
            if (pos < 0)
                return -1;

            if (int.TryParse(message.AsSpan(Prefix.Length, pos - Prefix.Length), NumberStyles.Integer, CultureInfo.InvariantCulture, out var i))
                return i;

            return -1;
        }
    }
}

namespace SheetReader
{
    public class XlsxBookFormat : BookFormat
    {
        public override BookFormatType Type => BookFormatType.Xlsx;
    }
}

namespace SheetReader.Utilities
{
    internal static class Extensions
    {
        public static bool EqualsIgnoreCase(this string? thisString, string? text, bool trim = false)
        {
            if (trim)
            {
                thisString = thisString.Nullify();
                text = text.Nullify();
            }

            if (thisString == null)
                return text == null;

            if (text == null)
                return false;

            if (thisString.Length != text.Length)
                return false;

            return string.Compare(thisString, text, StringComparison.OrdinalIgnoreCase) == 0;
        }

        public static string? Nullify(this string? text)
        {
            if (text == null)
                return null;

            if (string.IsNullOrWhiteSpace(text))
                return null;

            var t = text.Trim();
            return t.Length == 0 ? null : t;
        }

        [return: NotNullIfNotNull(nameof(path))]
        public static string? NormPath(string path) => path?.Replace(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }
}

