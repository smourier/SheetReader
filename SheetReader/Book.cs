using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using SheetReader.Utilities;
using Extensions = SheetReader.Utilities.Extensions;

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

            if (format is JsonBookFormat json)
                return EnumerateJsonSheets(stream, json);

            throw new NotSupportedException();
        }

        protected virtual JsonSheet CreateJsonSheet(JsonElement element, JsonBookFormat format) => new(element, format);
        protected virtual IEnumerable<Sheet> EnumerateJsonSheets(Stream stream, JsonBookFormat format)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(format);

            var root = JsonSerializer.Deserialize<JsonElement>(stream, format.SerializerOptions);
            format.RootElement = root;
            JsonElement? sheets = null;
            if (format.SheetsPropertyName == null)
            {
                if (root.ValueKind == JsonValueKind.Object &&
                    ((root.TryGetProperty("sheets", out var sheetElement) && sheetElement.ValueKind == JsonValueKind.Array) ||
                    (root.TryGetProperty("Sheets", out sheetElement) && sheetElement.ValueKind == JsonValueKind.Array)))
                {
                    sheets = sheetElement;
                }
            }
            else if (root.ValueKind == JsonValueKind.Object &&
                root.TryGetProperty(format.SheetsPropertyName, out JsonElement element) && element.ValueKind == JsonValueKind.Array)
            {
                sheets = element;
            }

            if (!sheets.HasValue)
            {
                if (root.ValueKind == JsonValueKind.Object)
                {
                    var props = enumerateArrayProperties(root).ToArray();
                    var count = 0;

                    if (processProperties(props))
                    {
                        foreach (var property in props)
                        {
                            var propSheet = CreateJsonSheet(property.Value, format);
                            propSheet.Name = property.Name;
                            if (propSheet != null)
                            {
                                yield return propSheet;
                                count++;
                            }
                        }
                    }
                    if (count != 0)
                        yield break;
                }

                var sheet = CreateJsonSheet(root, format);
                if (sheet != null)
                {
                    readSheet(root, sheet);
                    yield return sheet;
                }
                yield break;
            }

            foreach (var sheetElement in sheets.Value.EnumerateArray())
            {
                var sheet = CreateJsonSheet(sheetElement, format);
                if (sheet == null)
                    continue;

                readSheet(sheetElement, sheet);
                yield return sheet;
            }

            void readSheet(JsonElement element, Sheet sheet)
            {
                sheet.Name = element.GetNullifiedString("name") ?? element.GetNullifiedString("Name") ?? format.Name;
                if (element.GetNullableBoolean("isHidden").GetValueOrDefault() || element.GetNullableBoolean("IsHidden").GetValueOrDefault())
                {
                    sheet.IsVisible = false;
                }
            }

            bool processProperties(IReadOnlyList<JsonProperty> properties)
            {
                if (properties.Count <= 1)
                    return false;

                if (properties.Count == 2)
                {
                    var rowsName = Extensions.Nullify(format.RowsPropertyName) ?? "rows";
                    var columnsName = Extensions.Nullify(format.ColumnsPropertyName) ?? "columns";
                    if (Extensions.EqualsIgnoreCase(properties[0].Name, columnsName) && Extensions.EqualsIgnoreCase(properties[1].Name, rowsName))
                        return false;

                    if (Extensions.EqualsIgnoreCase(properties[0].Name, rowsName) && Extensions.EqualsIgnoreCase(properties[1].Name, columnsName))
                        return false;
                }
                return true;
            }

            static IEnumerable<JsonProperty> enumerateArrayProperties(JsonElement element)
            {
                if (element.ValueKind != JsonValueKind.Object)
                    yield break;

                foreach (var property in element.EnumerateObject())
                {
                    if (property.Value.ValueKind == JsonValueKind.Array)
                        yield return property;
                }
            }
        }

        public class JsonSheet : Sheet, IWithJsonElement
        {
            private readonly Dictionary<string, int> _columnsFromRows = new(StringComparer.OrdinalIgnoreCase);

            public JsonSheet(JsonElement element, JsonBookFormat format)
            {
                ArgumentNullException.ThrowIfNull(format);
                Element = element;
                Format = format;
            }

            public JsonElement Element { get; }
            public JsonBookFormat Format { get; }
            public IReadOnlyDictionary<string, int> ColumnsFromRows => _columnsFromRows;

            public virtual int AddOrGetColumnFromRows(string name)
            {
                ArgumentNullException.ThrowIfNull(name);
                if (!_columnsFromRows.TryGetValue(name, out var index))
                {
                    index = _columnsFromRows.Count;
                    _columnsFromRows.Add(name, index);
                }
                return index;
            }

            public override IEnumerable<Column> EnumerateColumns()
            {
                var count = 0;
                foreach (var item in EnumerateDeclaredColumns())
                {
                    yield return item;
                    count++;
                }

                if (count != 0 || _columnsFromRows.Count <= 0)
                {
                    yield break;
                }

                foreach (var kv in _columnsFromRows)
                {
                    yield return new Column { Index = kv.Value, Name = kv.Key };
                }
            }

            public virtual IEnumerable<Column> EnumerateDeclaredColumns()
            {
                JsonElement columns;
                if (Format.ColumnsPropertyName == null)
                {
                    if (Element.ValueKind != JsonValueKind.Object)
                        yield break;

                    if (Element.TryGetProperty("columns", out var element2) && element2.ValueKind == JsonValueKind.Array)
                    {
                        columns = element2;
                    }
                    else
                    {
                        if (!Element.TryGetProperty("Columns", out element2) || element2.ValueKind != JsonValueKind.Array)
                            yield break;

                        columns = element2;
                    }
                }
                else
                {
                    if (Element.ValueKind != JsonValueKind.Object || !Element.TryGetProperty(Format.ColumnsPropertyName, out var element) || element.ValueKind != JsonValueKind.Array)
                        yield break;

                    columns = element;
                }

                var index = 0;
                foreach (var columnElement in columns.EnumerateArray())
                {
                    Column column;
                    switch (columnElement.ValueKind)
                    {
                        case JsonValueKind.Object:
                            column = CreateColumn(columnElement);
                            if (column == null)
                                continue;

                            column.Name = columnElement.GetNullifiedString("name")
                                ?? columnElement.GetNullifiedString("value")
                                ?? columnElement.GetNullifiedString("Name")
                                ?? columnElement.GetNullifiedString("Value");
                            break;

                        case JsonValueKind.String:
                            column = CreateColumn();
                            if (column == null)
                                continue;

                            column.Name = columnElement.GetString();
                            break;

                        case JsonValueKind.Number:
                            column = CreateColumn();
                            if (column == null)
                                continue;

                            column.Name = columnElement.GetInt64().ToString();
                            break;

                        default:
                            column = CreateColumn();
                            if (column == null)
                                continue;

                            break;
                    }

                    if (column.Name != null)
                    {
                        column.Index = columnElement.GetNullableInt32("index") ?? columnElement.GetNullableInt32("Index").GetValueOrDefault(index);
                        yield return column;
                        index++;
                    }
                }
            }

            public override IEnumerable<Row> EnumerateRows()
            {
                JsonElement rows;
                if (Format.RowsPropertyName == null)
                {
                    if (Element.ValueKind == JsonValueKind.Array)
                    {
                        rows = Element;
                    }
                    else if (Element.TryGetProperty("rows", out var element) && element.ValueKind == JsonValueKind.Array)
                    {
                        rows = element;
                    }
                    else if (Element.TryGetProperty("Rows", out element) && element.ValueKind == JsonValueKind.Array)
                    {
                        rows = element;
                    }
                    else
                    {
                        var arrayElement = getFirstArrayProperty();
                        if (arrayElement != null)
                        {
                            rows = arrayElement.Value;
                        }
                        else
                            yield break;

                        JsonElement? getFirstArrayProperty()
                        {
                            if (Element.ValueKind != JsonValueKind.Object)
                                return null;

                            var first = Element.EnumerateObject().FirstOrDefault(p => p.Value.ValueKind == JsonValueKind.Array).Value;
                            if (first.ValueKind != JsonValueKind.Array)
                                return null;

                            return first;
                        }
                    }
                }
                else
                {
                    if (!Element.TryGetProperty(Format.RowsPropertyName, out var element) || element.ValueKind != JsonValueKind.Array)
                        yield break;

                    rows = element;
                }

                var index = 0;
                foreach (var rowElement in rows.EnumerateArray())
                {
                    var row = CreateRow(this, rowElement, Format);
                    if (row != null)
                    {
                        row.Index = index;
                        index++;
                        yield return row;
                    }
                }
            }

            protected virtual JsonColumn CreateColumn(JsonElement element) => new(element);
            protected virtual JsonRow CreateRow(JsonSheet sheet, JsonElement element, JsonBookFormat format) => new(sheet, element, format);
        }

        public class JsonColumn(JsonElement element) : Column, IWithJsonElement
        {
            public JsonElement Element { get; } = element;
        }

        public class JsonCell(JsonElement element) : Cell, IWithJsonElement
        {
            public JsonElement Element { get; } = element;
        }

        public class JsonRow : Row, IWithJsonElement
        {
            public JsonRow(JsonSheet sheet, JsonElement element, JsonBookFormat format)
            {
                ArgumentNullException.ThrowIfNull(sheet);
                ArgumentNullException.ThrowIfNull(format);
                Sheet = sheet;
                Element = element;
                Format = format;
            }

            public JsonSheet Sheet { get; }
            public JsonElement Element { get; }
            public JsonBookFormat Format { get; }

            public override IEnumerable<Cell> EnumerateCells()
            {
                if (Element.ValueKind == JsonValueKind.Array)
                {
                    var index = 0;
                    foreach (var cellElement2 in Element.EnumerateArray())
                    {
                        var cell = readCell(cellElement2);
                        cell.ColumnIndex = index;
                        yield return cell;
                        index++;
                    }
                }
                else
                {
                    if (Element.ValueKind != JsonValueKind.Object)
                        yield break;

                    if ((Element.TryGetProperty("cells", out var cellsElement) || Element.TryGetProperty("Cells", out cellsElement)) && cellsElement.ValueKind == JsonValueKind.Array)
                    {
                        var index = 0;
                        foreach (var cellElement in cellsElement.EnumerateArray())
                        {
                            var cell = readCell(cellElement);
                            cell.ColumnIndex = index;
                            yield return cell;
                            index++;
                        }
                        yield break;
                    }

                    foreach (var property in Element.EnumerateObject())
                    {
                        var index = Sheet.AddOrGetColumnFromRows(property.Name);
                        var cellElement = Element.GetProperty(property.Name);
                        var cell = readCell(cellElement);
                        cell.ColumnIndex = index;
                        yield return cell;
                    }
                }

                Cell readCell(JsonElement cellElement)
                {
                    Cell cell;
                    if (cellElement.ValueKind == JsonValueKind.Object)
                    {
                        cell = CreateJsonCell(cellElement);
                        if (cellElement.TryGetProperty("value", out var property) || cellElement.TryGetProperty("Value", out property))
                        {
                            cell.Value = Extensions.ToObject(property, Format.Options);
                            cell.RawValue = property.GetRawText();
                        }
                        else
                        {
                            cell.Value = Extensions.ToObject(cellElement, Format.Options);
                            cell.RawValue = cellElement.GetRawText();
                        }

                        if ((cellElement.TryGetProperty("isError", out property) || cellElement.TryGetProperty("IsError", out property)) && Extensions.ToObject(property, Format.Options) is bool error)
                        {
                            cell.IsError = error;
                        }
                    }
                    else
                    {
                        cell = CreateCell();
                        cell.Value = Extensions.ToObject(cellElement, Format.Options);
                        cell.RawValue = cellElement.GetRawText();
                    }
                    return cell;
                }
            }

            protected virtual Cell CreateCell() => new();
            protected virtual JsonCell CreateJsonCell(JsonElement element) => new(element);
        }

        protected virtual IEnumerable<Sheet> EnumerateCsvSheets(Stream stream, CsvBookFormat format)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(format);
            var sheet = CreateCsvSheet(stream, format);
            if (sheet != null)
                yield return sheet;

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

            var wb = CreateXlsxBook(format, archive, relPath, workBookDoc, workBookRelsDoc);
            if (wb != null)
            {
                foreach (var sheet in wb.EnumerateSheets())
                {
                    yield return sheet;
                }
            }
        }

        protected virtual XlsxBook CreateXlsxBook(XlsxBookFormat format, ZipArchive archive, string relativePath, XDocument workbookDocument, XDocument relsDocument) => new(format, archive, relativePath, workbookDocument, relsDocument);

        protected class XlsxBook
        {
            public XlsxBook(XlsxBookFormat format, ZipArchive archive, string relativePath, XDocument workbookDocument, XDocument relsDocument)
            {
                ArgumentNullException.ThrowIfNull(format);
                ArgumentNullException.ThrowIfNull(archive);
                ArgumentNullException.ThrowIfNull(relativePath);
                ArgumentNullException.ThrowIfNull(workbookDocument);
                ArgumentNullException.ThrowIfNull(relsDocument);
                Format = format;
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

            public XlsxBookFormat Format { get; }
            public ZipArchive Archive { get; }
            public string RelativePath { get; }
            public XDocument WorkbookDocument { get; }
            public XDocument RelsDocument { get; }
            public IReadOnlyList<string> SharedStrings { get; }
            public IReadOnlyDictionary<int, XlsxFormat> Formats { get; }

            protected virtual XlsxFormat CreateXlsxFormat(XElement element) => new(this, element);
            protected virtual XlsxSheet CreateXlsxSheet(XElement element, XmlReader reader) => new(Format, this, element, reader);

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
            public XlsxSheet(XlsxBookFormat format, XlsxBook book, XElement element, XmlReader reader)
            {
                ArgumentNullException.ThrowIfNull(format);
                ArgumentNullException.ThrowIfNull(book);
                ArgumentNullException.ThrowIfNull(element);
                ArgumentNullException.ThrowIfNull(reader);
                Format = format;
                Book = book;
                Element = element;
                Reader = reader;
                Name = element.Attribute("name")?.Value!;
                var state = element.Attribute("state")?.Value;
                if (Extensions.EqualsIgnoreCase(state, "hidden"))
                {
                    IsVisible = false;
                }
            }

            public XlsxBookFormat Format { get; }
            public XlsxBook Book { get; }
            public XElement Element { get; }
            public XmlReader Reader { get; }
            public IDictionary<int, Column> Columns { get; } = new Dictionary<int, Column>();

            public override IEnumerable<Column> EnumerateColumns() => Columns.Values.OrderBy(c => c.Index);

            public override IEnumerable<Row> EnumerateRows()
            {
                if (Format.LoadOptions.HasFlag(LoadOptions.FirstRowDefinesColumns))
                {
                    var count = 0;
                    foreach (var row in EnumerateDeclaredRows())
                    {
                        if (count == 0)
                        {
                            Columns.Clear();
                            foreach (var cell in row.EnumerateCells())
                            {
                                var index = Columns.Count;
                                var name = Extensions.Nullify(string.Format(CultureInfo.InvariantCulture, "{0}", cell.Value)) ?? index.ToString();
                                var col = new Column() { Index = index, Name = name };
                                Columns.Add(index, col);
                            }
                        }
                        else
                        {
                            yield return row;
                        }
                        count++;
                    }
                }
                else
                {
                    foreach (var row in EnumerateDeclaredRows())
                    {
                        yield return row;
                    }
                }
            }

            public virtual IEnumerable<Row> EnumerateDeclaredRows()
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
                                    if (hidden != null && Extensions.EqualsIgnoreCase(hidden, "true") || hidden == "1")
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

                if (!row.Sheet.Format.LoadOptions.HasFlag(LoadOptions.FirstRowDefinesColumns) &&
                    row.Sheet.Columns.Count <= ColumnIndex)
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
                Name = format.Name;
                Reader = new CsvReader(stream, format.AllowCharacterAmbiguity, format.LoadOptions.HasFlag(LoadOptions.FirstRowDefinesColumns), format.Quote, format.Separator, format.Encoding);
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
