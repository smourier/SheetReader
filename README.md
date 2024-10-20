# SheetReader
A simple CSV, XLSX or JSON data sheet reader. In the CSV and XLSX cases, it doesn't allocate memory for the whole input file (or stream) but reads rows and cells on demand.

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
    }
    Console.WriteLine("Visible rows:" + visibleRows);
    Console.WriteLine("Invisible rows:" + invisibleRows);

The reader code (not the WPF control and sample) is also available as a single .cs file: [SheetReader.cs](Amalgamation/SheetReader.cs)

There's also a read-only WPF control that allows to see what's been read by the **SheetReader**:

![image](https://github.com/smourier/SheetReader/assets/5328574/6c32c034-0703-4879-88b7-7a615bfffee1)

Supports keyboard navigation, selection and focus (mouse & keyboard):

![image](https://github.com/smourier/SheetReader/assets/5328574/0eca72a2-ff5f-46b0-9fda-6f8e404cfdf6)

Also supports column resizing, by mouse or programmatically.
Can be programmatically customized to use different styles (color, alignement, etc.):

![image](https://github.com/smourier/SheetReader/assets/5328574/dd2d5a2b-fb14-41a2-a116-9ab8d67ec4c4)

Also supports column sorting, by mouse double-click on column header or programmatically. Can be programmatically customized to use different comparisoo algorithms.

![image](https://github.com/user-attachments/assets/9b06dd43-1bbb-47d3-bb6d-748c91aec8fd)

Also supports column moving, by mouse drag on column header or programmatically.

![image](https://github.com/user-attachments/assets/cb1a7316-f8c7-4a4c-bf94-67994ca3049f)

It can also  export data back as .JSON or .CSV files with various options:

![image](https://github.com/user-attachments/assets/f38780da-5fb7-41ce-97f7-59b8717624de)

