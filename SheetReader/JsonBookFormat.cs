using System.Text.Json;

namespace SheetReader
{
    public class JsonBookFormat : BookFormat
    {
        public override BookFormatType Type => BookFormatType.Json;

        public virtual JsonBookOptions Options { get; set; } = JsonBookOptions.ParseDates;
        public virtual string? SheetsPropertyName { get; set; }
        public virtual string? ColumnsPropertyName { get; set; }
        public virtual string? RowsPropertyName { get; set; }
        public virtual JsonWriterOptions WriterOptions { get; set; }
        public virtual JsonSerializerOptions SerializerOptions { get; set; } = new JsonSerializerOptions
        {
            AllowTrailingCommas = true,
            ReadCommentHandling = JsonCommentHandling.Skip
        };
    }
}
