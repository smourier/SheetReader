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
                    Console.WriteLine(sheet + " visible:" + sheet.IsVisible);
                    Console.WriteLine(string.Join("\t", sheet.EnumerateColumns()));

                    foreach (var row in sheet.EnumerateRows())
                    {
                        Console.WriteLine("#" + row.Index);
                        if (!row.IsVisible)
                        {
                            Console.WriteLine("HIDDEN! " + string.Join("\t", row.EnumerateCells()));
                            continue;
                        }

                        Console.WriteLine(string.Join("\t", row.EnumerateCells()));
                    }
                }
                return;
            }
        }
    }
}
