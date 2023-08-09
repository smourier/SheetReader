using System;
using System.IO;
using System.Linq;
using System.Text;

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

            Console.OutputEncoding = Encoding.UTF8;
            foreach (var file in Directory.EnumerateFiles(".", "*.xlsx"))
            {
                Console.WriteLine(file);
                var book = new Book();
                var visibleRows = 0;
                var invisibleRows = 0;
                foreach (var sheet in book.EnumerateSheets(file))
                {
                    Console.WriteLine(sheet + " visible:" + sheet.IsVisible);
                    Console.WriteLine(string.Join("\t", sheet.EnumerateColumns()));

                    foreach (var row in sheet.EnumerateRows())
                    {
                        //Console.WriteLine("#" + row.Index);
                        if (!row.IsVisible)
                        {
                            invisibleRows++;
                            //Console.WriteLine("HIDDEN! " + string.Join("\t", row.EnumerateCells()));
                            continue;
                        }
                        visibleRows++;

                        var cells = row.EnumerateCells().ToList();
                        //Console.WriteLine(string.Join("\t", row.EnumerateCells()));
                    }
                    break;
                }
                Console.WriteLine("Visible rows:" + visibleRows);
                Console.WriteLine("Invisible rows:" + invisibleRows);
                //return;
            }
        }
    }
}
