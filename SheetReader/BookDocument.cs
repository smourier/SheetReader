using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;
using SheetReader.Utilities;

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
        public virtual bool IsThreadSafe => false;

        protected virtual Book CreateBook() => new();
        protected virtual BookDocumentSheet CreateSheet(Sheet sheet) => new(sheet);
        protected virtual IList<BookDocumentSheet> CreateSheets() => [];

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

        public virtual void Export(string filePath, ExportOptions options = ExportOptions.None, BookFormat? format = null)
        {
            ArgumentNullException.ThrowIfNull(filePath);
            format ??= BookFormat.GetFromFileExtension(Path.GetExtension(filePath));
            ArgumentNullException.ThrowIfNull(format);
            Utilities.Extensions.FileEnsureDirectory(filePath);
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

            var firstRowIndex = options.HasFlag(ExportOptions.StartFromFirstRow) ? sheet.FirstRowIndex.Value : 0;
            for (var rowIndex = firstRowIndex; rowIndex <= sheet.LastRowIndex.Value; rowIndex++)
            {
                if (options.HasFlag(ExportOptions.FirstRowIsHeader) && rowIndex == firstRowIndex)
                    continue;

                sheet.Rows.TryGetValue(rowIndex, out var ro);
                var firstColumnIndex = options.HasFlag(ExportOptions.StartFromFirstColumn) ? sheet.FirstColumnIndex.Value : 0;
                for (var columnIndex = firstColumnIndex; columnIndex <= sheet.LastColumnIndex.Value; columnIndex++)
                {
                    BookDocumentCell? cell = null;
                    ro?.Cells.TryGetValue(columnIndex, out cell);
                    var text = string.Format(CultureInfo.InvariantCulture, "{0}", cell?.Value);
                    Utilities.Extensions.WriteCsv(writer, text, columnIndex < sheet.LastColumnIndex);
                    if (columnIndex == sheet.LastColumnIndex)
                    {
                        writer.WriteLine();
                    }
                }
            }
        }

        protected virtual void ExportAsJson(Stream stream, ExportOptions options, JsonBookFormat format)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(format);

            var fo = format.WriterOptions;
            if (options.HasFlag(ExportOptions.JsonIndented))
            {
                fo.Indented = true;
            }

            using var writer = new Utf8JsonWriter(stream, fo);
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

            void writeCell(BookDocumentCell? cell)
            {
                if (cell != null && cell.IsError)
                {
                    writer.WriteStartObject();
                    writer.WritePropertyName("isError");
                    writer.WriteBooleanValue(value: true);
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

            void writeSheet(BookDocumentSheet sheet)
            {
                if (sheet.FirstColumnIndex.HasValue && sheet.LastColumnIndex.HasValue)
                {
                    if (sheet.ColumnsHaveBeenGenerated && !options.HasFlag(ExportOptions.JsonRowsAsObject))
                    {
                        writer.WriteStartArray();
                        if (sheet.FirstRowIndex.HasValue && sheet.LastRowIndex.HasValue)
                        {
                            var firstRowIndex = options.HasFlag(ExportOptions.StartFromFirstRow) ? sheet.FirstRowIndex.Value : 0;
                            for (var rowIndex = firstRowIndex; rowIndex <= sheet.LastRowIndex.Value; rowIndex++)
                            {
                                if (!options.HasFlag(ExportOptions.FirstRowIsHeader) || rowIndex != firstRowIndex)
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

                        Dictionary<int, string>? colNames = null;
                        if (!options.HasFlag(ExportOptions.JsonRowsAsObject))
                        {
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
                        else
                        {
                            colNames = [];
                            for (var columnIndex = sheet.FirstColumnIndex.Value; columnIndex <= sheet.LastColumnIndex.Value; columnIndex++)
                            {
                                string? name = null;
                                if (options.HasFlag(ExportOptions.FirstRowIsHeader) &&
                                    sheet.FirstRowIndex.HasValue && sheet.Rows.TryGetValue(sheet.FirstRowIndex.Value, out var row) &&
                                    row != null && row.Cells.TryGetValue(columnIndex, out var cell) &&
                                    cell != null && cell.Value != null)
                                {
                                    name = string.Format(CultureInfo.InvariantCulture, "{0}", cell.Value).Nullify();
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
                                if (!options.HasFlag(ExportOptions.FirstRowIsHeader) || rowIndex != firstRowIndex)
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
    }
}
