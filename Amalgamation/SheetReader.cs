//NOSONAR
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
AssemblyVersion: 3.2.1.0
AssemblyFileVersion: 3.2.1.0
*/
using Extensions = SheetReader.Utilities.Extensions;
using SheetReader.Utilities;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Compression;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Xml;
using System;

#pragma warning disable IDE0079 // Remove unnecessary suppression
#pragma warning disable IDE0130 // Namespace does not match folder structure

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
	
	        public virtual JsonSheet CreateJsonSheet(JsonElement element, JsonBookFormat format) => new(element, format);
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
	                    var cellsName = Extensions.Nullify(format.CellsPropertyName) ?? "cells";
	                    var rowsName = Extensions.Nullify(format.RowsPropertyName) ?? "rows";
	                    var columnsName = Extensions.Nullify(format.ColumnsPropertyName) ?? "columns";
	                    if (Extensions.EqualsIgnoreCase(properties[0].Name, columnsName) && Extensions.EqualsIgnoreCase(properties[1].Name, rowsName))
	                        return false;
	
	                    if (Extensions.EqualsIgnoreCase(properties[0].Name, rowsName) && Extensions.EqualsIgnoreCase(properties[1].Name, columnsName))
	                        return false;
	
	                    if (Extensions.EqualsIgnoreCase(properties[0].Name, columnsName) && Extensions.EqualsIgnoreCase(properties[1].Name, cellsName))
	                        return false;
	
	                    if (Extensions.EqualsIgnoreCase(properties[0].Name, cellsName) && Extensions.EqualsIgnoreCase(properties[1].Name, columnsName))
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
	            public int LastDefinedColumnIndex { get; protected set; } = -1;
	
	            public virtual void EnsureColumn(int columnIndex)
	            {
	                if (columnIndex > LastDefinedColumnIndex)
	                {
	                    LastDefinedColumnIndex = columnIndex;
	                }
	            }
	
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
	
	            public virtual Cell ReadCell(JsonElement cellElement)
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
	
	            public override IEnumerable<Column> EnumerateColumns()
	            {
	                var count = 0;
	                foreach (var item in EnumerateDeclaredColumns())
	                {
	                    yield return item;
	                    count++;
	                }
	
	                foreach (var kv in _columnsFromRows)
	                {
	                    yield return new Column { Index = kv.Value, Name = kv.Key };
	                    count++;
	                }
	
	                while (count <= LastDefinedColumnIndex)
	                {
	                    yield return new Column { Index = count, Name = count.ToString() };
	                    count++;
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
	                            column = CreateJsonColumn(columnElement);
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
	
	            public virtual IEnumerable<Row> EnumerateRowsFromCells(JsonElement cells)
	            {
	                if (cells.ValueKind != JsonValueKind.Array)
	                    return [];
	
	                var rows = new Dictionary<int, JsonCellsRow>();
	                foreach (var cellElement in cells.EnumerateArray())
	                {
	                    var rowIndex = cellElement.GetNullableInt32("r") ?? 0;
	                    if (!rows.TryGetValue(rowIndex, out var row))
	                    {
	                        row = CreateJsonCellsRow(this, Format);
	                        row.Index = rowIndex;
	                        rows.Add(rowIndex, row);
	                    }
	
	                    var colIndex = cellElement.GetNullableInt32("c") ?? 0;
	                    var cell = ReadCell(cellElement);
	                    row.SetCell(colIndex, cell);
	                }
	                return rows.OrderBy(r => r.Key).Select(r => r.Value);
	            }
	
	            public override IEnumerable<Row> EnumerateRows()
	            {
	                JsonElement? cells = null;
	                if (Format.CellsPropertyName == null)
	                {
	                    if (Element.TryGetProperty("cells", out var element) && element.ValueKind == JsonValueKind.Array)
	                    {
	                        cells = element;
	                    }
	                    else if (Element.TryGetProperty("Cells", out element) && element.ValueKind == JsonValueKind.Array)
	                    {
	                        cells = element;
	                    }
	                }
	                else
	                {
	                    if (Element.TryGetProperty(Format.CellsPropertyName, out var element) && element.ValueKind == JsonValueKind.Array)
	                    {
	                        cells = element;
	                    }
	                }
	
	                if (cells.HasValue)
	                {
	                    foreach (var cell in EnumerateRowsFromCells(cells.Value))
	                    {
	                        yield return cell;
	                    }
	                    yield break;
	                }
	
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
	                    var row = CreateJsonRow(this, rowElement, Format);
	                    if (row != null)
	                    {
	                        row.Index = index;
	                        index++;
	                        yield return row;
	                    }
	                }
	            }
	
	            public virtual JsonColumn CreateJsonColumn(JsonElement element) => new(element);
	            public virtual JsonRow CreateJsonRow(JsonSheet sheet, JsonElement element, JsonBookFormat format) => new(sheet, element, format);
	            public virtual JsonCellsRow CreateJsonCellsRow(JsonSheet sheet, JsonBookFormat format) => new(sheet, format);
	            public virtual JsonCell CreateJsonCell(JsonElement element) => new(element);
	            public virtual Cell CreateCell() => new();
	        }
	
	        public class JsonCellsRow : Row
	        {
	            private readonly Dictionary<int, Cell> _cells = [];
	            private int _max = -1;
	
	            public JsonCellsRow(JsonSheet sheet, JsonBookFormat format)
	            {
	                ArgumentNullException.ThrowIfNull(sheet);
	                ArgumentNullException.ThrowIfNull(format);
	                Sheet = sheet;
	                Format = format;
	            }
	
	            public JsonSheet Sheet { get; }
	            public JsonBookFormat Format { get; }
	            public IReadOnlyDictionary<int, Cell> Cells => _cells;
	
	            public virtual void SetCell(int columnIndex, Cell cell)
	            {
	                _cells[columnIndex] = cell;
	                if (columnIndex > _max)
	                {
	                    _max = columnIndex;
	                }
	
	                Sheet.EnsureColumn(columnIndex);
	            }
	
	            public override IEnumerable<Cell> EnumerateCells()
	            {
	                if (_max < 0)
	                    yield break;
	
	                for (var i = 0; i <= _max; i++)
	                {
	                    if (Cells.TryGetValue(i, out var cell))
	                    {
	                        cell.ColumnIndex = i;
	                        yield return cell;
	                    }
	                    else
	                    {
	                        var nullCell = Sheet.CreateCell();
	                        nullCell.ColumnIndex = i;
	                        yield return nullCell;
	                    }
	                }
	            }
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
	                        var cell = Sheet.ReadCell(cellElement2);
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
	                            var cell = Sheet.ReadCell(cellElement);
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
	                        var cell = Sheet.ReadCell(cellElement);
	                        cell.ColumnIndex = index;
	                        yield return cell;
	                    }
	                }
	            }
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
	
	public class BookDocument
	    {
	        private readonly IList<BookDocumentSheet> _sheets;
	
	        public event EventHandler<StateChangedEventArgs>? StateChanged;
	
	        public BookDocument()
	        {
	            _sheets = CreateSheets();
	            if (_sheets == null)
	                throw new InvalidOperationException();
	        }
	
	        public IList<BookDocumentSheet> Sheets => _sheets;
	        protected virtual Book CreateBook() => new();
	        protected virtual BookDocumentSheet CreateSheet(Sheet sheet) => new(this, sheet);
	        protected virtual IList<BookDocumentSheet> CreateSheets() => [];
	
	        protected virtual internal void OnStateChanged(object sender, StateChangedEventArgs e) => StateChanged?.Invoke(this, e);
	
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
	                        docSheet.Load(this, sheet);
	                        _sheets.Add(docSheet);
	                    }
	                }
	            }
	        }
	
	        public virtual void Export(string filePath, ExportOptions options = ExportOptions.None, BookFormat? format = null)
	        {
	            ArgumentNullException.ThrowIfNull(filePath);
	            format ??= BookFormat.GetFromFileExtension(Path.GetExtension(filePath));
	            ArgumentNullException.ThrowIfNull(format);
	            Extensions.FileEnsureDirectory(filePath);
	            using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
	            format.IsStreamOwned = true;
	            format.InputFilePath = filePath;
	            Export(stream, options, format);
	        }
	
	        public virtual void Export(Stream stream, ExportOptions options, BookFormat format)
	        {
	            ArgumentNullException.ThrowIfNull(stream);
	            ArgumentNullException.ThrowIfNull(format);
	            if (format is JsonBookFormat json)
	            {
	                ExportAsJson(stream, options, json);
	                return;
	            }
	
	            if (format is CsvBookFormat csv)
	            {
	                ExportAsCsv(stream, options, csv);
	                return;
	            }
	
	            throw new NotSupportedException();
	        }
	
	        protected virtual void ExportAsCsv(Stream stream, ExportOptions options, CsvBookFormat format)
	        {
	            ArgumentNullException.ThrowIfNull(stream);
	            ArgumentNullException.ThrowIfNull(format);
	
	            using var writer = new StreamWriter(stream, Encoding.Unicode);
	            if (Sheets.Count <= 0)
	                return;
	
	            var sheet = Sheets[0];
	            if (!sheet.FirstColumnIndex.HasValue || !sheet.LastColumnIndex.HasValue || !sheet.FirstRowIndex.HasValue || !sheet.LastRowIndex.HasValue)
	                return;
	
	
	            if (options.HasFlag(ExportOptions.CsvWriteColumns))
	            {
	                var firstColumnIndex = options.HasFlag(ExportOptions.StartFromFirstColumn) ? sheet.FirstColumnIndex.Value : 0;
	                for (var columnIndex = firstColumnIndex; columnIndex <= sheet.LastColumnIndex.Value; columnIndex++)
	                {
	                    sheet.Columns.TryGetValue(columnIndex, out var column);
	                    var name = column?.Name ?? columnIndex.ToString();
	                    Extensions.WriteCsv(writer, name, columnIndex < sheet.LastColumnIndex);
	                    if (columnIndex == sheet.LastColumnIndex)
	                    {
	                        writer.WriteLine();
	                    }
	                }
	            }
	
	            var firstRowIndex = options.HasFlag(ExportOptions.StartFromFirstRow) ? sheet.FirstRowIndex.Value : 0;
	            for (var rowIndex = firstRowIndex; rowIndex <= sheet.LastRowIndex.Value; rowIndex++)
	            {
	                if (options.HasFlag(ExportOptions.FirstRowDefinesColumns) && rowIndex == firstRowIndex)
	                    continue;
	
	                sheet.Rows.TryGetValue(rowIndex, out var ro);
	                var firstColumnIndex = options.HasFlag(ExportOptions.StartFromFirstColumn) ? sheet.FirstColumnIndex.Value : 0;
	                for (var columnIndex = firstColumnIndex; columnIndex <= sheet.LastColumnIndex.Value; columnIndex++)
	                {
	                    BookDocumentCell? cell = null;
	                    ro?.Cells.TryGetValue(columnIndex, out cell);
	                    var text = sheet.FormatValue(cell?.Value) ?? string.Empty;
	                    Extensions.WriteCsv(writer, text, columnIndex < sheet.LastColumnIndex);
	                    if (columnIndex == sheet.LastColumnIndex)
	                    {
	                        writer.WriteLine();
	                    }
	                }
	            }
	        }
	
	        public virtual void ExportAsJson(Stream stream, ExportOptions options, JsonBookFormat format)
	        {
	            ArgumentNullException.ThrowIfNull(stream);
	            ArgumentNullException.ThrowIfNull(format);
	
	            var fo = format.WriterOptions;
	            if (options.HasFlag(ExportOptions.JsonIndented))
	            {
	                fo.Indented = true;
	            }
	
	            using var writer = new Utf8JsonWriter(stream, fo);
	            WriteTo(writer, options, format);
	        }
	
	        public virtual void WriteTo(Utf8JsonWriter writer, ExportOptions options, JsonBookFormat format)
	        {
	            ArgumentNullException.ThrowIfNull(writer);
	            ArgumentNullException.ThrowIfNull(format);
	
	            var fo = format.WriterOptions;
	            if (options.HasFlag(ExportOptions.JsonIndented))
	            {
	                fo.Indented = true;
	            }
	
	            if (Sheets.Count > 0)
	            {
	                if (Sheets.Count == 1)
	                {
	                    writeSheet(Sheets[0]);
	                }
	                else
	                {
	                    writer.WriteStartObject();
	                    writer.WritePropertyName(format.SheetsPropertyName ?? "sheets");
	                    writer.WriteStartArray();
	                    foreach (var sheet in Sheets)
	                    {
	                        writeSheet(sheet);
	                    }
	                    writer.WriteEndArray();
	                    writer.WriteEndObject();
	                }
	            }
	            writer.Flush();
	
	            static bool isDefaultJsonValue(object? value)
	            {
	                if (value == null || Convert.IsDBNull(value))
	                    return true;
	
	                if (value is bool b && !b)
	                    return true;
	
	                if (value is int i && i == 0)
	                    return true;
	
	                if (value is long j && j == 0)
	                    return true;
	
	                if (value is float fl && fl == 0f)
	                    return true;
	
	                if (value is double db && db == 0.0)
	                    return true;
	
	                if (value is decimal dec && dec == 0m)
	                    return true;
	
	                if (value is uint ui && ui == 0)
	                    return true;
	
	                if (value is uint ul && ul == 0)
	                    return true;
	
	                return false;
	            }
	
	            void writePositionedCell(BookDocumentCell? cell, int rowIndex, int columnIndex)
	            {
	                // don't output null values
	                if (cell == null || cell.Value == null || Convert.IsDBNull(cell.Value))
	                    return;
	
	                writer.WriteStartObject();
	                writer.WriteNumber("r", rowIndex);
	                writer.WriteNumber("c", columnIndex);
	                writer.WritePropertyName("value");
	
	                if (cell.IsError)
	                {
	                    if (cell.Value != null)
	                    {
	                        writeCell(new BookDocumentCell(new Cell { Value = cell.Value }));
	                    }
	
	                    writer.WriteBoolean("isError", true);
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                if (cell.Value is string s)
	                {
	                    writer.WriteStringValue(s);
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                if (cell.Value is bool b)
	                {
	                    writer.WriteBooleanValue(b);
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                if (cell.Value is int i32)
	                {
	                    writer.WriteNumberValue(i32);
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                if (cell.Value is long l)
	                {
	                    writer.WriteNumberValue(l);
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                if (cell.Value is uint ui)
	                {
	                    writer.WriteNumberValue(ui);
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                if (cell.Value is ulong ul)
	                {
	                    writer.WriteNumberValue(ul);
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                if (cell.Value is float flt)
	                {
	                    writer.WriteNumberValue(flt);
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                if (cell.Value is double dbl)
	                {
	                    writer.WriteNumberValue(dbl);
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                if (cell.Value is decimal dec)
	                {
	                    writer.WriteNumberValue(dec);
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                if (cell.Value is byte[] bytes)
	                {
	                    writer.WriteBase64StringValue(bytes);
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                if (cell.Value is IDictionary dictionary)
	                {
	                    foreach (DictionaryEntry kv in dictionary)
	                    {
	                        writer.WritePropertyName(kv.Key.ToString()!);
	                        writeCell(new BookDocumentCell(new Cell() { Value = kv.Value }));
	                    }
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                if (cell.Value is Array array && array.Rank == 1)
	                {
	                    writer.WriteStartArray();
	                    for (var i = 0; i < array.Length; i++)
	                    {
	                        var item = array.GetValue(i);
	                        writeCell(new BookDocumentCell(new Cell() { Value = item }));
	                    }
	                    writer.WriteEndArray();
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                s = string.Format(CultureInfo.InvariantCulture, "{0}", cell.Value);
	                writer.WriteStringValue(s);
	                writer.WriteEndObject();
	            }
	
	            void writeCell(BookDocumentCell? cell)
	            {
	                if (cell != null && cell.IsError)
	                {
	                    writer.WriteStartObject();
	                    writer.WriteBoolean("isError", true);
	                    if (cell.Value != null)
	                    {
	                        writer.WritePropertyName("value");
	                        writeCell(new BookDocumentCell(new Cell { Value = cell.Value }));
	                    }
	
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                if (cell == null || cell.Value == null || Convert.IsDBNull(cell.Value))
	                {
	                    writer.WriteNullValue();
	                    return;
	                }
	
	                if (cell.Value is string s)
	                {
	                    writer.WriteStringValue(s);
	                    return;
	                }
	
	                if (cell.Value is bool b)
	                {
	                    writer.WriteBooleanValue(b);
	                    return;
	                }
	
	                if (cell.Value is int i32)
	                {
	                    writer.WriteNumberValue(i32);
	                    return;
	                }
	
	                if (cell.Value is long l)
	                {
	                    writer.WriteNumberValue(l);
	                    return;
	                }
	
	                if (cell.Value is uint ui)
	                {
	                    writer.WriteNumberValue(ui);
	                    return;
	                }
	
	                if (cell.Value is ulong ul)
	                {
	                    writer.WriteNumberValue(ul);
	                    return;
	                }
	
	                if (cell.Value is float flt)
	                {
	                    writer.WriteNumberValue(flt);
	                    return;
	                }
	
	                if (cell.Value is double dbl)
	                {
	                    writer.WriteNumberValue(dbl);
	                    return;
	                }
	
	                if (cell.Value is decimal dec)
	                {
	                    writer.WriteNumberValue(dec);
	                    return;
	                }
	
	                if (cell.Value is byte[] bytes)
	                {
	                    writer.WriteBase64StringValue(bytes);
	                    return;
	                }
	
	                if (cell.Value is IDictionary dictionary)
	                {
	                    writer.WriteStartObject();
	                    foreach (DictionaryEntry kv in dictionary)
	                    {
	                        writer.WritePropertyName(kv.Key.ToString()!);
	                        writeCell(new BookDocumentCell(new Cell() { Value = kv.Value }));
	                    }
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                if (cell.Value is Array array && array.Rank == 1)
	                {
	                    writer.WriteStartArray();
	                    for (var i = 0; i < array.Length; i++)
	                    {
	                        var item = array.GetValue(i);
	                        writeCell(new BookDocumentCell(new Cell() { Value = item }));
	                    }
	                    writer.WriteEndArray();
	                    return;
	                }
	
	                s = string.Format(CultureInfo.InvariantCulture, "{0}", cell.Value);
	                writer.WriteStringValue(s);
	            }
	
	            void writeRow(BookDocumentSheet sheet, BookDocumentRow? row)
	            {
	                writer.WriteStartArray();
	                if (sheet.FirstColumnIndex.HasValue && sheet.LastColumnIndex.HasValue)
	                {
	                    var firstColumnIndex = options.HasFlag(ExportOptions.StartFromFirstColumn) ? sheet.FirstColumnIndex.Value : 0;
	                    for (var columnIndex = firstColumnIndex; columnIndex <= sheet.LastColumnIndex.Value; columnIndex++)
	                    {
	                        BookDocumentCell? cell = null;
	                        row?.Cells.TryGetValue(columnIndex, out cell);
	                        writeCell(cell);
	                    }
	                }
	                writer.WriteEndArray();
	            }
	
	            void writeRowAsObjects(BookDocumentSheet sheet, BookDocumentRow? row, Dictionary<int, string> colNames)
	            {
	                writer.WriteStartObject();
	                if (sheet.FirstColumnIndex.HasValue && sheet.LastColumnIndex.HasValue)
	                {
	                    for (var columnIndex = sheet.FirstColumnIndex.Value; columnIndex <= sheet.LastColumnIndex.Value; columnIndex++)
	                    {
	                        BookDocumentCell? cell = null;
	                        row?.Cells.TryGetValue(columnIndex, out cell);
	                        if (!options.HasFlag(ExportOptions.JsonNoDefaultCellValues) || (cell != null && !isDefaultJsonValue(cell.Value)))
	                        {
	                            string colName = colNames[columnIndex];
	                            writer.WritePropertyName(colName);
	                            writeCell(cell);
	                        }
	                    }
	                }
	                writer.WriteEndObject();
	            }
	
	            void writeColumns(BookDocumentSheet sheet)
	            {
	                if (!sheet.FirstColumnIndex.HasValue || !sheet.LastColumnIndex.HasValue)
	                    return;
	
	                writer.WritePropertyName(format.ColumnsPropertyName ?? "columns");
	                writer.WriteStartArray();
	                var firstColumnIndex = options.HasFlag(ExportOptions.StartFromFirstColumn) ? sheet.FirstColumnIndex.Value : 0;
	                for (var columnIndex = firstColumnIndex; columnIndex <= sheet.LastColumnIndex.Value; columnIndex++)
	                {
	                    sheet.Columns.TryGetValue(columnIndex, out var col);
	                    writer.WriteStringValue(col?.Name);
	                }
	                writer.WriteEndArray();
	            }
	
	            void writeSheetHeader(BookDocumentSheet sheet)
	            {
	                if (!string.IsNullOrEmpty(sheet.Name))
	                {
	                    writer.WritePropertyName("name");
	                    writer.WriteStringValue(sheet.Name);
	                }
	
	                if (sheet.IsHidden)
	                {
	                    writer.WritePropertyName("isHidden");
	                    writer.WriteBooleanValue(value: true);
	                }
	            }
	
	            void writeSheet(BookDocumentSheet sheet)
	            {
	                if (!sheet.FirstColumnIndex.HasValue || !sheet.LastColumnIndex.HasValue)
	                    return;
	
	                if (options.HasFlag(ExportOptions.JsonCellByCell))
	                {
	                    writer.WriteStartObject();
	                    writeSheetHeader(sheet);
	
	                    if (!sheet.ColumnsHaveBeenGenerated)
	                    {
	                        writeColumns(sheet);
	                    }
	
	                    writer.WritePropertyName(format.RowsPropertyName ?? "cells");
	                    writer.WriteStartArray();
	
	                    if (sheet.FirstRowIndex.HasValue && sheet.LastRowIndex.HasValue)
	                    {
	                        var rowOffset = sheet.ColumnsHaveBeenGenerated ? 0 : -1;
	                        for (var rowIndex = sheet.FirstRowIndex.Value; rowIndex <= sheet.LastRowIndex.Value; rowIndex++)
	                        {
	                            sheet.Rows.TryGetValue(rowIndex, out var row);
	                            if (row != null)
	                            {
	                                for (var columnIndex = sheet.FirstColumnIndex.Value; columnIndex <= sheet.LastColumnIndex.Value; columnIndex++)
	                                {
	                                    if (row.Cells.TryGetValue(columnIndex, out var cell))
	                                    {
	                                        writePositionedCell(cell, rowIndex + rowOffset, columnIndex);
	                                    }
	                                }
	                            }
	                        }
	                    }
	
	                    writer.WriteEndArray();
	                    writer.WriteEndObject();
	                    return;
	                }
	
	                if (sheet.ColumnsHaveBeenGenerated && !options.HasFlag(ExportOptions.JsonRowsAsObject))
	                {
	                    writer.WriteStartArray();
	                    if (sheet.FirstRowIndex.HasValue && sheet.LastRowIndex.HasValue)
	                    {
	                        var firstRowIndex = options.HasFlag(ExportOptions.StartFromFirstRow) ? sheet.FirstRowIndex.Value : 0;
	                        for (var rowIndex = firstRowIndex; rowIndex <= sheet.LastRowIndex.Value; rowIndex++)
	                        {
	                            if (!options.HasFlag(ExportOptions.FirstRowDefinesColumns) || rowIndex != firstRowIndex)
	                            {
	                                sheet.Rows.TryGetValue(rowIndex, out var row2);
	                                writeRow(sheet, row2);
	                            }
	                        }
	                    }
	                    writer.WriteEndArray();
	                }
	                else
	                {
	                    writer.WriteStartObject();
	                    writeSheetHeader(sheet);
	
	                    Dictionary<int, string>? colNames = null;
	                    if (!options.HasFlag(ExportOptions.JsonRowsAsObject))
	                    {
	                        writeColumns(sheet);
	                    }
	                    else
	                    {
	                        colNames = [];
	                        for (var columnIndex = sheet.FirstColumnIndex.Value; columnIndex <= sheet.LastColumnIndex.Value; columnIndex++)
	                        {
	                            string? name = null;
	                            if (options.HasFlag(ExportOptions.FirstRowDefinesColumns) &&
	                                sheet.FirstRowIndex.HasValue && sheet.Rows.TryGetValue(sheet.FirstRowIndex.Value, out var row) &&
	                                row != null && row.Cells.TryGetValue(columnIndex, out var cell) &&
	                                cell != null && cell.Value != null)
	                            {
	                                name = Extensions.Nullify(string.Format(CultureInfo.InvariantCulture, "{0}", cell.Value));
	                            }
	
	                            if (name == null)
	                            {
	                                sheet.Columns.TryGetValue(columnIndex, out var col);
	                                name = col?.Name ?? columnIndex.ToString();
	                            }
	                            colNames[columnIndex] = name;
	                        }
	                    }
	
	                    writer.WritePropertyName(format.RowsPropertyName ?? "rows");
	                    writer.WriteStartArray();
	                    if (sheet.FirstRowIndex.HasValue && sheet.LastRowIndex.HasValue)
	                    {
	                        var firstRowIndex = options.HasFlag(ExportOptions.StartFromFirstRow) ? sheet.FirstRowIndex.Value : 0;
	                        for (var rowIndex = firstRowIndex; rowIndex <= sheet.LastRowIndex.Value; rowIndex++)
	                        {
	                            if (!options.HasFlag(ExportOptions.FirstRowDefinesColumns) || rowIndex != firstRowIndex)
	                            {
	                                sheet.Rows.TryGetValue(rowIndex, out var row);
	                                if (options.HasFlag(ExportOptions.JsonRowsAsObject))
	                                {
	                                    writeRowAsObjects(sheet, row, colNames!);
	                                }
	                                else
	                                {
	                                    writeRow(sheet, row);
	                                }
	                            }
	                        }
	                    }
	                    writer.WriteEndArray();
	                    writer.WriteEndObject();
	                }
	            }
	        }
	    }
	
	public class BookDocumentCell : IWithValue
	    {
	        public BookDocumentCell(Cell cell)
	        {
	            ArgumentNullException.ThrowIfNull(cell);
	            Value = cell.Value;
	        }
	
	        public virtual object? Value { get; set; }
	        public virtual bool IsError => this is BookDocumentCellError;
	
	        public override string ToString() => Value?.ToString() ?? string.Empty;
	    }
	
	public class BookDocumentCellError(Cell cell)
	        : BookDocumentCell(cell)
	    {
	    }
	
	public class BookDocumentJsonCell : BookDocumentCell, IWithJsonElement
	    {
	        public BookDocumentJsonCell(Book.JsonCell cell)
	            : base(cell)
	        {
	            ArgumentNullException.ThrowIfNull(cell);
	            Value = cell.Value;
	            Element = cell.Element;
	        }
	
	        public JsonElement Element { get; }
	    }
	
	public class BookDocumentJsonCellError(Book.JsonCell cell)
	        : BookDocumentCellError(cell)
	    {
	    }
	
	public class BookDocumentRow
	    {
	        private readonly IDictionary<int, BookDocumentCell> _cells;
	
	        public BookDocumentRow(BookDocument book, BookDocumentSheet sheet, Row row)
	        {
	            ArgumentNullException.ThrowIfNull(sheet);
	            ArgumentNullException.ThrowIfNull(book);
	            ArgumentNullException.ThrowIfNull(row);
	            _cells = CreateCells();
	            if (_cells == null)
	                throw new InvalidOperationException();
	
	            Row = row;
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
	
	                var e = new StateChangedEventArgs(StateChangedType.CellAdded, sheet, this, null, bdCell);
	                book.OnStateChanged(this, e);
	                if (e.Cancel)
	                    break;
	            }
	        }
	
	        public Row Row { get; }
	        public int RowIndex { get; }
	        public virtual bool IsHidden { get; }
	        public int? FirstCellIndex { get; }
	        public int? LastCellIndex { get; }
	        public IDictionary<int, BookDocumentCell> Cells => _cells;
	
	        protected virtual IDictionary<int, BookDocumentCell> CreateCells() => new Dictionary<int, BookDocumentCell>();
	        protected virtual BookDocumentCell CreateCell(Cell cell)
	        {
	            ArgumentNullException.ThrowIfNull(cell);
	            if (cell is Book.JsonCell json)
	                return cell.IsError ? new BookDocumentJsonCellError(json) : new BookDocumentJsonCell(json);
	
	            return cell.IsError ? new BookDocumentCellError(cell) : new BookDocumentCell(cell);
	        }
	
	        public override string ToString() => RowIndex.ToString();
	    }
	
	public class BookDocumentSheet
	    {
	        private readonly IDictionary<int, BookDocumentRow> _rows;
	        private readonly IDictionary<int, Column> _columns;
	
	        public BookDocumentSheet(BookDocument book, Sheet sheet)
	        {
	            ArgumentNullException.ThrowIfNull(book);
	            ArgumentNullException.ThrowIfNull(sheet);
	            _rows = CreateRows();
	            _columns = CreateColumns();
	            if (_rows == null || _columns == null)
	                throw new InvalidOperationException();
	
	            Name = sheet.Name ?? string.Empty;
	            IsHidden = !sheet.IsVisible;
	
	            var e = new StateChangedEventArgs(StateChangedType.SheetAdded, this);
	            book.OnStateChanged(this, e);
	            if (e.Cancel)
	                return;
	        }
	
	        public virtual void Load(BookDocument book, Sheet sheet)
	        {
	            ArgumentNullException.ThrowIfNull(book);
	            ArgumentNullException.ThrowIfNull(sheet);
	
	            StateChangedEventArgs e;
	            foreach (var row in sheet.EnumerateRows())
	            {
	                var rowData = CreateRow(book, row);
	                if (rowData == null)
	                    continue;
	
	                _rows[row.Index] = rowData;
	
	                if (LastRowIndex == null || row.Index > LastRowIndex)
	                {
	                    LastRowIndex = row.Index;
	                }
	
	                if (FirstRowIndex == null || row.Index < FirstRowIndex)
	                {
	                    FirstRowIndex = row.Index;
	                }
	
	                e = new StateChangedEventArgs(StateChangedType.RowAdded, this, rowData);
	                book.OnStateChanged(this, e);
	                if (e.Cancel)
	                    break;
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
	
	                e = new StateChangedEventArgs(StateChangedType.ColumnAddded, this, null, col);
	                book.OnStateChanged(this, e);
	                if (e.Cancel)
	                    break;
	            }
	
	            if (_columns.Count == 0 && _rows.Count > 0)
	            {
	                ColumnsHaveBeenGenerated = true;
	                for (var i = 0; i < _rows[0].Cells.Count; i++)
	                {
	                    var col = new Column { Index = i };
	                    _columns[i] = col;
	                    if (!LastColumnIndex.HasValue || col.Index > LastColumnIndex)
	                    {
	                        LastColumnIndex = col.Index;
	                    }
	
	                    if (!FirstColumnIndex.HasValue || col.Index < FirstColumnIndex)
	                    {
	                        FirstColumnIndex = col.Index;
	                    }
	
	                    e = new StateChangedEventArgs(StateChangedType.ColumnAddded, this, null, col);
	                    book.OnStateChanged(this, e);
	                    if (e.Cancel)
	                        break;
	                }
	            }
	        }
	
	        public virtual string Name { get; }
	        public virtual bool IsHidden { get; }
	        public bool ColumnsHaveBeenGenerated { get; protected set; }
	        public int? FirstColumnIndex { get; protected set; }
	        public int? LastColumnIndex { get; protected set; }
	        public int? FirstRowIndex { get; protected set; }
	        public int? LastRowIndex { get; protected set; }
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
	
	        public virtual string? FormatValue(object? value)
	        {
	            if (value is null)
	                return null;
	
	            if (value is string s)
	                return s;
	
	            if (value is IDictionary dictionary)
	            {
	                var sb = new StringBuilder("{");
	                var first = true;
	                foreach (DictionaryEntry kv in dictionary)
	                {
	                    if (first)
	                    {
	                        first = !first;
	                    }
	                    else
	                    {
	                        sb.Append(", ");
	                    }
	
	                    sb.Append(kv.Key);
	                    sb.Append('=');
	                    sb.Append(FormatValue(kv.Value));
	                }
	                sb.Append('}');
	                return sb.ToString();
	            }
	
	            if (value is Array array && array.Rank == 1)
	            {
	                var sb = new StringBuilder("[");
	                var first = true;
	                for (var i = 0; i < array.Length; i++)
	                {
	                    if (first)
	                    {
	                        first = !first;
	                    }
	                    else
	                    {
	                        sb.Append(", ");
	                    }
	
	                    sb.Append(FormatValue(array.GetValue(i)));
	                }
	                sb.Append(']');
	                return sb.ToString();
	            }
	
	            return string.Format(CultureInfo.CurrentCulture, "{0}", value);
	        }
	
	        protected virtual BookDocumentRow CreateRow(BookDocument book, Row row) => new(book, this, row);
	        protected virtual IDictionary<int, Column> CreateColumns() => new Dictionary<int, Column>();
	        protected virtual IDictionary<int, BookDocumentRow> CreateRows() => new Dictionary<int, BookDocumentRow>();
	    }
	
	public abstract class BookFormat
	    {
	        private string? _name;
	
	        protected BookFormat()
	        {
	        }
	
	        public abstract BookFormatType Type { get; }
	        public virtual bool IsStreamOwned { get; set; }
	        public virtual string? InputFilePath { get; set; }
	        public virtual LoadOptions LoadOptions { get; set; }
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
	
	public enum BookFormatType
	    {
	        Automatic,
	        Csv,
	        Xlsx,
	        Json,
	    }
	
	public class Cell : IWithValue
	    {
	        public virtual int ColumnIndex { get; set; }
	        public virtual object? Value { get; set; }
	        public virtual bool IsError { get; set; }
	        public virtual string? RawValue { get; set; }
	
	        public override string ToString() => Value?.ToString() ?? string.Empty;
	    }
	
	public class Column : IWithValue
	    {
	        public virtual int Index { get; set; }
	        public virtual string? Name { get; set; }
	        object? IWithValue.Value => Name;
	
	        public override string ToString() => Extensions.Nullify(Name) ?? Index.ToString();
	    }
	
	public class CsvBookFormat : BookFormat
	    {
	        public override BookFormatType Type => BookFormatType.Csv;
	
	        public virtual bool AllowCharacterAmbiguity { get; set; } = false;
	        public virtual char Quote { get; set; } = '"';
	        public virtual char Separator { get; set; } = ';';
	        public virtual Encoding? Encoding { get; set; }
	    }
	
	public class CsvReader : IDisposable
	    {
	        private readonly List<string> _columns = [];
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
	
	[Flags]
	    public enum ExportOptions
	    {
	        None = 0x0,
	        StartFromFirstColumn = 0x1,
	        StartFromFirstRow = 0x2,
	        FirstRowDefinesColumns = 0x4,
	
	        // json only
	        JsonRowsAsObject = 0x8,
	        JsonNoDefaultCellValues = 0x10,
	        JsonIndented = 0x20,
	        JsonCellByCell = 0x80,
	
	        // csv only
	        CsvWriteColumns = 0x40,
	    }
	
	public interface IWithJsonElement
	    {
	        JsonElement Element { get; }
	    }
	
	public interface IWithValue
	    {
	        object? Value { get; }
	    }
	
	public class JsonBookFormat : BookFormat
	    {
	        public override BookFormatType Type => BookFormatType.Json;
	        public static IReadOnlyList<string> WellKnownRootPropertyNames { get; } = ["sheets", "rows", "columns", "cells", "name"];
	        public static IReadOnlyList<string> WellKnownRowsPropertyNames { get; } = ["cells"];
	        public static IReadOnlyList<string> WellKnownColumnPropertyNames { get; } = ["name", "value", "index"];
	        public static IReadOnlyList<string> WellKnownCellPropertyNames { get; } = ["value", "isError"];
	
	        public virtual JsonBookOptions Options { get; set; } = JsonBookOptions.ParseDates;
	        public virtual string? SheetsPropertyName { get; set; }
	        public virtual string? ColumnsPropertyName { get; set; }
	        public virtual string? RowsPropertyName { get; set; }
	        public virtual string? CellsPropertyName { get; set; }
	        public virtual JsonWriterOptions WriterOptions { get; set; }
	        public virtual JsonElement RootElement { get; set; }
	        public virtual JsonSerializerOptions SerializerOptions { get; set; } = new JsonSerializerOptions
	        {
	            AllowTrailingCommas = true,
	            ReadCommentHandling = JsonCommentHandling.Skip
	        };
	    }
	
	[Flags]
	    public enum JsonBookOptions
	    {
	        None = 0x0,
	        ParseDates = 0x1,
	
	        Default = ParseDates,
	    }
	
	[Flags]
	    public enum LoadOptions
	    {
	        None = 0x0,
	        FirstRowDefinesColumns = 0x1, // sometimes, this can be guessed
	    }
	
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
	
	public class RowCol : IEquatable<RowCol>
	    {
	        public RowCol()
	        {
	        }
	
	        public RowCol(int rowIndex, int columnIndex)
	        {
	            RowIndex = rowIndex;
	            ColumnIndex = columnIndex;
	        }
	
	        public virtual int RowIndex { get; set; }
	        public virtual int ColumnIndex { get; set; }
	
	        public string ExcelReference => Row.GetExcelColumnName(ColumnIndex) + (RowIndex + 1).ToString();
	
	        public override string ToString() => RowIndex + "," + ColumnIndex;
	        public override bool Equals(object? obj) => Equals(obj as RowCol);
	        public bool Equals(RowCol? other) => other is not null && RowIndex == other.RowIndex && ColumnIndex == other.ColumnIndex;
	        public override int GetHashCode() => RowIndex.GetHashCode() ^ ColumnIndex.GetHashCode();
	        public static bool operator !=(RowCol? obj1, RowCol? obj2) => !(obj1 == obj2);
	        public static bool operator ==(RowCol? obj1, RowCol? obj2)
	        {
	            if (ReferenceEquals(obj1, obj2))
	                return true;
	
	            if (obj1 is null)
	                return false;
	
	            if (obj2 is null)
	                return false;
	
	            return obj1.Equals(obj2);
	        }
	    }
	
	public abstract class Sheet
	    {
	        public virtual string? Name { get; set; }
	        public virtual bool IsVisible { get; set; } = true;
	
	        public abstract IEnumerable<Column> EnumerateColumns();
	        public abstract IEnumerable<Row> EnumerateRows();
	
	        protected internal virtual Column CreateColumn() => new();
	
	        public override string ToString() => Name ?? string.Empty;
	    }
	
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
	
	public enum StateChangedType
	    {
	        SheetAdded,
	        ColumnAddded,
	        RowAdded,
	        CellAdded,
	    }
	
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
	
	        public static bool FileEnsureDirectory(string path, bool throwOnError = true)
	        {
	            ArgumentNullException.ThrowIfNull(path);
	            if (!Path.IsPathRooted(path))
	            {
	                path = Path.GetFullPath(path);
	            }
	
	            var dir = Path.GetDirectoryName(path);
	            if (dir == null)
	                return false;
	
	            if (Directory.Exists(dir))
	                return true;
	
	            try
	            {
	                Directory.CreateDirectory(dir);
	                return true;
	            }
	            catch
	            {
	                if (throwOnError)
	                    throw;
	
	                return false;
	            }
	        }
	
	        public static bool? GetNullableBoolean(this JsonElement element, string propertyName)
	        {
	            ArgumentNullException.ThrowIfNull(propertyName);
	            if (element.ValueKind != JsonValueKind.Object)
	                return null;
	
	            if (!element.TryGetProperty(propertyName, out var property))
	                return null;
	
	            return property.ValueKind switch
	            {
	                JsonValueKind.False => false,
	                JsonValueKind.True => true,
	                _ => null,
	            };
	        }
	
	        public static string? GetNullifiedString(this JsonElement element, string propertyName)
	        {
	            ArgumentNullException.ThrowIfNull(propertyName);
	            if (element.ValueKind != JsonValueKind.Object)
	                return null;
	
	            if (!element.TryGetProperty(propertyName, out var property))
	                return null;
	
	            if (property.ValueKind == JsonValueKind.String)
	                return property.GetString();
	
	            return property.GetRawText();
	        }
	
	        public static int? GetNullableInt32(this JsonElement element, string propertyName)
	        {
	            ArgumentNullException.ThrowIfNull(propertyName);
	            if (element.ValueKind != JsonValueKind.Object)
	                return null;
	
	            if (!element.TryGetProperty(propertyName, out var property))
	                return null;
	
	            if (property.ValueKind == JsonValueKind.Number)
	            {
	                if (property.TryGetInt32(out var i))
	                    return i;
	
	                return null;
	            }
	
	            var text = property.GetRawText();
	            if (int.TryParse(text, out var value))
	                return value;
	
	            return null;
	        }
	
	        public static object? ToObject(this JsonElement element, JsonBookOptions options)
	        {
	            element.TryConvertToObject(options, out var value);
	            return value;
	        }
	
	        public static bool TryConvertToObject(this JsonElement element, JsonBookOptions options, out object? value)
	        {
	            switch (element.ValueKind)
	            {
	                case JsonValueKind.Null:
	                    value = null;
	                    return true;
	
	                case JsonValueKind.Object:
	                    var dic = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
	                    foreach (var child in element.EnumerateObject())
	                    {
	                        if (!child.Value.TryConvertToObject(options, out var childValue))
	                        {
	                            value = null;
	                            return false;
	                        }
	
	                        dic[child.Name] = childValue;
	                    }
	
	                    if (dic.Count == 0)
	                    {
	                        value = null;
	                        return true;
	                    }
	
	                    value = dic;
	                    return true;
	
	                case JsonValueKind.Array:
	                    var objects = new object?[element.GetArrayLength()];
	                    var i = 0;
	                    foreach (var child in element.EnumerateArray())
	                    {
	                        if (!child.TryConvertToObject(options, out var childValue2))
	                        {
	                            value = null;
	                            return false;
	                        }
	                        objects[i++] = childValue2;
	                    }
	                    value = objects;
	                    return true;
	
	                case JsonValueKind.String:
	                    var str = element.ToString();
	                    if (options.HasFlag(JsonBookOptions.ParseDates) && DateTime.TryParseExact(str, ["o", "r", "s"], null, DateTimeStyles.None, out var dt))
	                    {
	                        value = dt;
	                        return true;
	                    }
	                    value = str;
	                    return true;
	
	                case JsonValueKind.Number:
	                    if (element.TryGetInt32(out var i2))
	                    {
	                        value = i2;
	                        return true;
	                    }
	
	                    if (element.TryGetInt64(out var i64))
	                    {
	                        value = i64;
	                        return true;
	                    }
	
	                    if (element.TryGetDecimal(out var dec))
	                    {
	                        value = dec;
	                        return true;
	                    }
	
	                    if (element.TryGetDouble(out var dbl))
	                    {
	                        value = dbl;
	                        return true;
	                    }
	                    break;
	
	                case JsonValueKind.True:
	                    value = true;
	                    return true;
	
	                case JsonValueKind.False:
	                    value = false;
	                    return true;
	            }
	
	            value = null;
	            return false;
	        }
	
	        public static void WriteCsv(TextWriter writer, string cell, bool addSeparator = true, bool forExcel = true)
	        {
	            ArgumentNullException.ThrowIfNull(writer);
	
	            var max = 32758;
	            if (forExcel && cell != null && cell.Length > max)
	            {
	                cell = cell[..max];
	            }
	
	            if (cell != null && cell.IndexOfAny(['\t', '\r', '\n', '"']) >= 0)
	            {
	                writer.Write('"');
	                writer.Write(cell.Replace("\"", "\"\""));
	                writer.Write('"');
	            }
	            else
	            {
	                writer.Write(cell);
	            }
	
	            if (addSeparator)
	            {
	                writer.Write('\t');
	            }
	        }
	    }
}

#pragma warning restore IDE0130 // Namespace does not match folder structure
#pragma warning restore IDE0079 // Remove unnecessary suppression
