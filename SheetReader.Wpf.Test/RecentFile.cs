using System;

namespace SheetReader.Wpf.Test
{
    public class RecentFile
    {
        public string? FilePath { get; set; }
        public DateTime LastAccessTime { get; set; } = DateTime.Now;
        public LoadOptions LoadOptions { get; set; }

        public override string ToString() => LastAccessTime + " " + FilePath;
    }
}
