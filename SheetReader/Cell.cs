namespace SheetReader
{
    public class Cell
    {
        public virtual object? Value { get; set; }

        public override string ToString() => Value?.ToString() ?? string.Empty;
    }
}
