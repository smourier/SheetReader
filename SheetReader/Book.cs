using System;
using System.Collections.Generic;
using System.IO;
using SheetReader.Utilities;

namespace SheetReader
{
    public class Book
    {
        public virtual IEnumerable<Sheet> EnumerateSheets(string filePath, BookFormat? format = null)
        {
            ArgumentNullException.ThrowIfNull(filePath);

            if (format == null)
            {
                var ext = Path.GetExtension(filePath);
                if (ext.EqualsIgnoreCase(".csv"))
                {
                    format = new CsvBookFormat();
                }
                else if (ext.EqualsIgnoreCase(".xlsx"))
                {
                    format = new XlsxBookFormat();
                }
                else
                {
                    ArgumentNullException.ThrowIfNull(format);
                }
            }

            var stream = File.OpenRead(filePath);
            format.IsStreamOwned = true;
            format.InputFilePath = filePath;
            return EnumerateSheets(stream, format);
        }

        public virtual IEnumerable<Sheet> EnumerateSheets(Stream stream, BookFormat format)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(format);

            if (format is CsvBookFormat csv)
                return EnumerateCsvSheets(stream, csv);

            if (format is XlsxBookFormat xlsx)
                return EnumerateXlsxSheets(stream, xlsx);

            throw new NotSupportedException();
        }

        protected virtual IEnumerable<Sheet> EnumerateCsvSheets(Stream stream, CsvBookFormat format)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(format);
            yield return new CsvSheet(stream, format);
            if (format.IsStreamOwned)
            {
                stream.Dispose();
            }
        }

        protected virtual IEnumerable<Sheet> EnumerateXlsxSheets(Stream stream, XlsxBookFormat format)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(format);
            yield break;
        }

        protected class CsvRow : Row
        {
            public CsvRow(IEnumerator<string> cells)
            {
                ArgumentNullException.ThrowIfNull(cells);
                Cells = cells;
            }

            public IEnumerator<string> Cells { get; }

            public override IEnumerable<Cell> EnumerateCells()
            {
                while (Cells.MoveNext())
                {
                    yield return new Cell { Value = Cells.Current };
                }
            }
        }

        protected class CsvSheet : Sheet, IDisposable
        {
            private bool _disposedValue;

            public CsvSheet(Stream stream, CsvBookFormat format)
            {
                ArgumentNullException.ThrowIfNull(stream);
                ArgumentNullException.ThrowIfNull(format);
                if (format.InputFilePath != null)
                {
                    Name = Path.GetFileNameWithoutExtension(format.InputFilePath);
                }
                Reader = new CsvReader(stream, format.AllowCharacterAmbiguity, format.ReadHeaderRow, format.Quote, format.Separator, format.Encoding);
            }

            public CsvReader Reader { get; }

            public override IEnumerable<Row> EnumerateRows()
            {
                var rowIndex = 0;
                while (!Reader.EndOfStream)
                {
                    var cells = Reader.ReadRow().GetEnumerator();
                    yield return new CsvRow(cells) { Index = rowIndex };
                    rowIndex++;
                }
            }

            public override IEnumerable<Column> EnumerateColumns()
            {
                for (var i = 0; i < Reader.Columns.Count; i++)
                {
                    yield return new Column { Name = Reader.Columns[i], Index = i };
                }
            }

            protected virtual void Dispose(bool disposing)
            {
                if (!_disposedValue)
                {
                    if (disposing)
                    {
                        // dispose managed state (managed objects)
                        Reader.Dispose();
                    }

                    // free unmanaged resources (unmanaged objects) and override finalizer
                    // set large fields to null
                    _disposedValue = true;
                }
            }

            // // override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
            // ~CsvSheet()
            // {
            //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            //     Dispose(disposing: false);
            // }

            public void Dispose()
            {
                // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
                Dispose(disposing: true);
                GC.SuppressFinalize(this);
            }
        }
    }
}
