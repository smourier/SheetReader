using System;
using System.IO;

namespace SheetReader.Test
{
    class Program
    {
        public static void Main()
        {
            //foreach (var file in Directory.EnumerateFiles(".", "*.csv"))
            //{
            //    Console.WriteLine(file);
            //    var book = new Book();
            //    foreach (var sheet in book.EnumerateSheets(file))
            //    {
            //        Console.WriteLine(sheet);
            //        Console.WriteLine(string.Join("\t", sheet.EnumerateColumns()));

            //        foreach (var row in sheet.EnumerateRows())
            //        {
            //            Console.WriteLine(string.Join("\t", row.EnumerateCells()));
            //        }
            //    }
            //}

            foreach (var file in Directory.EnumerateFiles(".", "*.xlsx"))
            {
                Console.WriteLine(file);
                var book = new Book();
                foreach (var sheet in book.EnumerateSheets(file))
                {
                    Console.WriteLine(sheet);
                    Console.WriteLine(string.Join("\t", sheet.EnumerateColumns()));

                    foreach (var row in sheet.EnumerateRows())
                    {
                        Console.WriteLine(string.Join("\t", row.EnumerateCells()));
                    }
                }
            }
        }
    }
}
