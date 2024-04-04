using System;

namespace SheetReader.Wpf.Test
{
    public class RecentFile
    {
        public string? FilePath { get; set; }
        public DateTime LastAccessTime { get; set; } = DateTime.Now;

        public override string ToString() => LastAccessTime + " " + FilePath;
    }
}
