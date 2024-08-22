using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Text;

namespace SheetReader
{
    public class CsvReader : IDisposable
    {
        private readonly List<string> _columns = [];
        private Stream? _ownedStream;
        private bool _disposedValue;

        public CsvReader(string filePath, bool allowCharacterAmbiguity, bool readHeaderRow, char quote, char separator, Encoding? encoding)
        {
            ArgumentNullException.ThrowIfNull(filePath);

            var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            _ownedStream = stream;
            Initialize(stream, allowCharacterAmbiguity, readHeaderRow, quote, separator, encoding);
        }

        public CsvReader(Stream stream, bool allowCharacterAmbiguity, bool readHeaderRow, char quote, char separator, Encoding? encoding)
        {
            Initialize(stream, allowCharacterAmbiguity, readHeaderRow, quote, separator, encoding);
        }

        public CsvReader(TextReader reader, bool allowCharacterAmbiguity, bool readHeaderRow, char quote, char separator)
        {
            ArgumentNullException.ThrowIfNull(reader);
            BaseReader = reader;
            InitializeReader(allowCharacterAmbiguity, readHeaderRow, quote, separator);
        }

        public TextReader? BaseReader { get; private set; }
        public IReadOnlyList<string> Columns => _columns;
        public virtual char Separator { get; protected set; }
        public virtual char Quote { get; protected set; }
        public virtual Encoding? Encoding { get; protected set; }
        public virtual int LineNumber { get; protected set; }
        public virtual int ColumnNumber { get; protected set; }
        public virtual int RowNumber { get; protected set; }
        public virtual bool ReadHeaderRow { get; protected set; }
        protected bool IsReaderOwned { get; set; }

        public bool EndOfStream
        {
            get
            {
                if (BaseReader == null)
                    return true;

                if (BaseReader is StreamReader sr)
                    return sr.EndOfStream;

                return BaseReader.Peek() < 0;
            }
        }

        public static IEnumerable<string> ReadLine(string line) => ReadLine(line, '"', ';');
        public static IEnumerable<string> ReadLine(string line, char quote, char separator)
        {
            if (line == null)
                yield break;

            using var reader = new StringReader(line);
            using var csvReader = new CsvReader(reader, true, false, quote, separator);
            foreach (string cell in csvReader.ReadRow())
            {
                yield return cell;
            }
        }

        [MemberNotNull(nameof(BaseReader), nameof(Encoding))]
        protected void Initialize(Stream stream, bool allowCharacterAmbiguity, bool readHeaderRow, char quote, char separator, Encoding? encoding)
        {
            ArgumentNullException.ThrowIfNull(stream);
            IsReaderOwned = true;
            Encoding = encoding ?? Encoding.UTF8;
            BaseReader = new StreamReader(stream, Encoding, true);
            InitializeReader(allowCharacterAmbiguity, readHeaderRow, quote, separator);
        }

        protected void InitializeReader(bool allowCharacterAmbiguity, bool readHeaderRow, char quote, char separator)
        {
            Quote = quote;
            Separator = separator;
            BaseReader!.Peek(); // force encoding detection
            DetermineFormat(allowCharacterAmbiguity);
            ReadHeaderRow = readHeaderRow;
            if (readHeaderRow)
            {
                ReadHeader();
            }
        }

        public virtual IEnumerable<string> ReadRow()
        {
            if (BaseReader == null)
                yield break;

            var inQuote = false;
            int next;
            char c;
            var cell = new StringBuilder();
            var hasCell = true;
            ColumnNumber = 0;
            do
            {
                var i = BaseReader.Read();
                if (i < 0)
                {
                    if (hasCell)
                        yield return cell.ToString();

                    break;
                }

                ColumnNumber++;
                next = BaseReader.Peek();

                c = (char)i;
                if (inQuote)
                {
                    if (c == '\n')
                    {
                        LineNumber++;
                    }
                    else if (c == '\r')
                    {
                        if (next == '\n')
                        {
                            hasCell = true;
                            cell.Append(c);
                            ColumnNumber++;
                            c = (char)BaseReader.Read();
                        }
                        LineNumber++;
                    }

                    if (c == Quote)
                    {
                        if (next == Quote)
                        {
                            BaseReader.Read();
                            hasCell = true;
                            cell.Append(Quote);
                        }
                        else
                        {
                            inQuote = false;
                        }
                        continue;
                    }

                    hasCell = true;
                    cell.Append(c);
                }
                else
                {
                    if (c == '\n')
                    {
                        if (hasCell)
                        {
                            yield return cell.ToString();
                            cell.Length = 0;
                        }
                        LineNumber++;
                        break;
                    }

                    if (c == '\r')
                    {
                        if (next == '\n')
                        {
                            BaseReader.Read();
                            ColumnNumber++;
                        }

                        if (hasCell)
                        {
                            yield return cell.ToString();
                            cell.Length = 0;
                        }
                        LineNumber++;
                        break;
                    }

                    if (c == Quote)
                    {
                        if (next == Quote)
                        {
                            BaseReader.Read();
                            hasCell = true;
                            cell.Append(Quote);
                        }
                        else
                        {
                            inQuote = true;
                        }
                        continue;
                    }

                    if (c == Separator)
                    {
                        if (hasCell)
                        {
                            yield return cell.ToString();
                            cell.Length = 0;
                        }
                        // note: we keep hasCell to true
                        continue;
                    }

                    hasCell = true;
                    cell.Append(c);
                }
            }
            while (true);
            RowNumber++;
        }

        protected virtual void ReadHeader()
        {
            if (BaseReader == null)
                throw new InvalidOperationException();

            foreach (var column in ReadRow())
            {
                _columns.Add(column);
            }
        }

        protected virtual void DetermineFormat(bool allowCharacterAmbiguity)
        {
            if (BaseReader == null)
                return;

            var sr = BaseReader as StreamReader;
            if (sr != null && sr.CurrentEncoding is UTF8Encoding)
                return;

            if (sr != null && sr.CurrentEncoding is UnicodeEncoding)
            {
                Separator = '\t';
                return;
            }

            if (!allowCharacterAmbiguity && sr != null)
                throw new SheetReaderException($"0002: CSV has ambiguous file encoding: '{sr.CurrentEncoding.WebName}'.");
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    // dispose managed state (managed objects)
                    var br = BaseReader;
                    BaseReader = null;
                    if (IsReaderOwned)
                    {
                        br?.Dispose();
                    }

                    _ownedStream?.Dispose();
                    _ownedStream = null;
                }

                _disposedValue = true;
            }
        }

        public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
    }
}
