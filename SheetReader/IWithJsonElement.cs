using System.Text.Json;

namespace SheetReader
{
    public interface IWithJsonElement
    {
        JsonElement Element { get; }
    }
}
