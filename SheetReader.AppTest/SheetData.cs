using System.Collections.Generic;

namespace SheetReader.AppTest
{
    public class SheetData
    {
        public string? Name { get; set; }
        public bool IsHidden { get; set; }
        public List<List<Cell>> Rows { get; } = new List<List<Cell>>();

        public override string ToString() => Name ?? string.Empty;
    }
}
