namespace System;

internal static class StringCompatibilityExtensions
{
    public static bool Contains(this string value, char comparisonValue, StringComparison comparisonType)
    {
        if (value is null)
        {
            throw new ArgumentNullException(nameof(value));
        }

        return value.IndexOf(comparisonValue.ToString(), comparisonType) >= 0;
    }

    public static bool Contains(this string value, string comparisonValue, StringComparison comparisonType)
    {
        if (value is null)
        {
            throw new ArgumentNullException(nameof(value));
        }

        return value.IndexOf(comparisonValue, comparisonType) >= 0;
    }

    public static string Replace(this string value, string oldValue, string? newValue, StringComparison comparisonType)
    {
        if (value is null)
        {
            throw new ArgumentNullException(nameof(value));
        }

        if (oldValue is null)
        {
            throw new ArgumentNullException(nameof(oldValue));
        }

        if (oldValue.Length == 0)
        {
            throw new ArgumentException("Old value cannot be empty.", nameof(oldValue));
        }

        var replacement = newValue ?? string.Empty;
        var builder = new System.Text.StringBuilder();
        var startIndex = 0;
        while (true)
        {
            var matchIndex = value.IndexOf(oldValue, startIndex, comparisonType);
            if (matchIndex < 0)
            {
                builder.Append(value, startIndex, value.Length - startIndex);
                return builder.ToString();
            }

            builder.Append(value, startIndex, matchIndex - startIndex);
            builder.Append(replacement);
            startIndex = matchIndex + oldValue.Length;
        }
    }
}
