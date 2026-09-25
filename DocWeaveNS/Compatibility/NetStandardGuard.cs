namespace DocWeave;

internal static class NetStandardGuard
{
    public static void ThrowIfNull<T>(T? value, string paramName) where T : class
    {
        if (value is null)
        {
            throw new ArgumentNullException(paramName);
        }
    }

    public static void ThrowIfNullOrWhiteSpace(string? value, string paramName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException("Value cannot be null or whitespace.", paramName);
        }
    }

    public static void ThrowIfLessThan<T>(T value, T other, string paramName) where T : IComparable<T>
    {
        if (value.CompareTo(other) < 0)
        {
            throw new ArgumentOutOfRangeException(paramName, value, $"Value must be greater than or equal to {other}.");
        }
    }
}
