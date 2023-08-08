using System.Text;

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
