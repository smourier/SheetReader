using System.Collections.Generic;
using System.Text.Json;

namespace SheetReader
{
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
}
