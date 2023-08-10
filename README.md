# SheetReader
A simple CSV or XLSX data sheet reader. It doesn't allocate memory for the whole file but reads rows and cells on demand.

Example:

    var book = new Book();
    var visibleRows = 0;
    var invisibleRows = 0;
    foreach (var sheet in book.EnumerateSheets(@"c:\path\blah.xlsx")) // or .csv
    {
        Console.WriteLine(sheet + " (visible:" + sheet.IsVisible + ")");
        foreach (var row in sheet.EnumerateRows())
        {
            if (!row.IsVisible)
            {
                invisibleRows++;
                continue;
            }
            
            visibleRows++;
            Console.WriteLine(string.Join("\t", row.EnumerateCells()));
        }
        break;
    }
    Console.WriteLine("Visible rows:" + visibleRows);
    Console.WriteLine("Invisible rows:" + invisibleRows);

The code is also available as a single .cs file: [SheetReader.cs](Amalgamation/SheetReader.cs)

