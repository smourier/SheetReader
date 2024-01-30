using System;
using System.Globalization;

namespace SheetReader
{
    [Serializable]
    public class SheetReaderException : Exception
    {
        public const string Prefix = "SHR";

        public SheetReaderException()
            : base(Prefix + "0001: SheetReader exception.")
        {
        }

        public SheetReaderException(string message)
            : base(Prefix + ":" + message)
        {
        }

        public SheetReaderException(Exception innerException)
            : base(null, innerException)
        {
        }

        public SheetReaderException(string message, Exception innerException)
            : base(Prefix + ":" + message, innerException)
        {
        }

        public int Code => GetCode(Message);

        public static int GetCode(string message)
        {
            if (message == null)
                return -1;

            if (!message.StartsWith(Prefix, StringComparison.Ordinal))
                return -1;

            var pos = message.IndexOf(':', Prefix.Length);
            if (pos < 0)
                return -1;

            if (int.TryParse(message.AsSpan(Prefix.Length, pos - Prefix.Length), NumberStyles.Integer, CultureInfo.InvariantCulture, out var i))
                return i;

            return -1;
        }
    }
}
