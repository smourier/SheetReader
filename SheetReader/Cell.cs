namespace SheetReader
{
    public class Cell
    {
        public virtual object? Value { get; set; }
        public virtual bool IsError { get; set; }
        public virtual string? RawValue { get; set; }

        public override string ToString() => Value?.ToString() ?? string.Empty;
    }
}
