using System.Globalization;

namespace DocWeave;

/// <summary>
/// Turns the text of a cell into a value of a property's type. Numbers and dates use the invariant culture, so the result does
/// not depend on the machine's locale.
/// </summary>
internal static class ValueConversion
{
    /// <summary>
    /// True when text can be converted to <paramref name="type"/>.
    /// </summary>
    public static bool IsSupported(Type type)
    {
        var target = Nullable.GetUnderlyingType(type) ?? type;
        return target == typeof(string)
            || target == typeof(bool)
            || target == typeof(decimal) || target == typeof(double) || target == typeof(float)
            || IsWholeNumber(target)
            || target == typeof(DateTime) || IsDateOnly(target) || IsTimeOnly(target)
            || target == typeof(DateTimeOffset) || target == typeof(TimeSpan)
            || target == typeof(Guid)
            || target.IsEnum;
    }

    /// <summary>
    /// True when a blank cell is acceptable for <paramref name="type"/>: any reference type, and any nullable value type.
    /// </summary>
    public static bool AllowsBlank(Type type)
    {
        return !type.IsValueType || Nullable.GetUnderlyingType(type) is not null;
    }

    /// <summary>
    /// Converts non-blank text.
    /// </summary>
    /// <param name="text">The cell text.</param>
    /// <param name="type">The property type. A nullable type is treated as its underlying type.</param>
    /// <param name="excelSerialDates">
    /// When true, a number is accepted for a date or time property and read as an Excel date serial (days since 1899-12-30,
    /// with the time as a fraction of a day). Excel stores dates that way, and the import returns the stored number.
    /// </param>
    /// <param name="value">The converted value.</param>
    /// <param name="error">Why the text could not be converted.</param>
    public static bool TryConvert(string text, Type type, bool excelSerialDates, out object? value, out string? error)
    {
        var target = Nullable.GetUnderlyingType(type) ?? type;
        value = null;
        error = null;

        if (target == typeof(string))
        {
            value = text;
            return true;
        }

        var trimmed = text.Trim();

        if (target == typeof(bool))
        {
            switch (trimmed.ToLowerInvariant())
            {
                case "true" or "1" or "yes" or "y":
                    value = true;
                    return true;
                case "false" or "0" or "no" or "n":
                    value = false;
                    return true;
                default:
                    return Fail(text, "true/false, yes/no or 1/0", out error);
            }
        }

        if (IsWholeNumber(target))
        {
            // Read as a decimal first so "3.0" (how a spreadsheet may store 3) is accepted, but "3.5" is not.
            if (!decimal.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || decimal.Truncate(number) != number)
            {
                return Fail(text, "a whole number", out error);
            }

            try
            {
                value = Convert.ChangeType(number, target, CultureInfo.InvariantCulture);
                return true;
            }
            catch (OverflowException)
            {
                error = $"'{text}' is outside the range of {target.Name}.";
                return false;
            }
        }

        if (target == typeof(decimal))
        {
            if (decimal.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            {
                value = number;
                return true;
            }

            return Fail(text, "a number", out error);
        }

        if (target == typeof(double) || target == typeof(float))
        {
            if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            {
                value = target == typeof(float) ? (float)number : number;
                return true;
            }

            return Fail(text, "a number", out error);
        }

        if (target == typeof(DateTime) || IsDateOnly(target) || target == typeof(DateTimeOffset))
        {
            if (TryParseDateTime(trimmed, excelSerialDates, out var date))
            {
                if (IsDateOnly(target))
                {
                    value = CreateDateOnly(target, date);
                }
                else if (target == typeof(DateTimeOffset))
                {
                    value = new DateTimeOffset(date);
                }
                else
                {
                    value = date;
                }

                return true;
            }

            return Fail(text, "a date", out error);
        }

        if (IsTimeOnly(target) || target == typeof(TimeSpan))
        {
            if (excelSerialDates && double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var days))
            {
                var span = TimeSpan.FromDays(days);
                value = IsTimeOnly(target) ? CreateTimeOnly(target, TimeSpan.FromTicks(span.Ticks % TimeSpan.TicksPerDay)) : span;
                return true;
            }

            if (IsTimeOnly(target) && TryCreateTimeOnly(target, trimmed, out value))
            {
                return true;
            }

            if (target == typeof(TimeSpan) && TimeSpan.TryParse(trimmed, CultureInfo.InvariantCulture, out var interval))
            {
                value = interval;
                return true;
            }

            return Fail(text, "a time", out error);
        }

        if (target == typeof(Guid))
        {
            if (Guid.TryParse(trimmed, out var guid))
            {
                value = guid;
                return true;
            }

            return Fail(text, "a GUID", out error);
        }

        if (target.IsEnum)
        {
            // A number is accepted only when the enum defines it. Enum.Parse would accept any number.
            try
            {
                var member = Enum.Parse(target, trimmed, true);
                if (Enum.IsDefined(target, member))
                {
                    value = member;
                    return true;
                }
            }
            catch (ArgumentException)
            {
            }

            error = $"'{text}' is not one of: {string.Join(", ", Enum.GetNames(target))}.";
            return false;
        }
        error = $"Type {target.Name} is not supported.";
        return false;
    }

    private static bool TryParseDateTime(string text, bool excelSerialDates, out DateTime date)
    {
        if (excelSerialDates && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial))
        {
            try
            {
                date = DateTime.FromOADate(serial);
                return true;
            }
            catch (ArgumentException)
            {
                date = default;
                return false;
            }
        }

        return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind | DateTimeStyles.AllowWhiteSpaces, out date);
    }

    private static bool IsWholeNumber(Type type)
    {
        return type == typeof(int) || type == typeof(long) || type == typeof(short) || type == typeof(byte)
            || type == typeof(uint) || type == typeof(ulong) || type == typeof(ushort) || type == typeof(sbyte);
    }

    private static bool IsDateOnly(Type type)
    {
        return type.FullName == "System.DateOnly";
    }

    private static bool IsTimeOnly(Type type)
    {
        return type.FullName == "System.TimeOnly";
    }

    private static object CreateDateOnly(Type type, DateTime date)
    {
        var method = type.GetMethod("FromDateTime", new[] { typeof(DateTime) })
            ?? throw new NotSupportedException("DateOnly is not available on this runtime.");
        return method.Invoke(null, new object[] { date })
            ?? throw new InvalidOperationException("DateOnly conversion returned null.");
    }

    private static object CreateTimeOnly(Type type, TimeSpan time)
    {
        var method = type.GetMethod("FromTimeSpan", new[] { typeof(TimeSpan) })
            ?? throw new NotSupportedException("TimeOnly is not available on this runtime.");
        return method.Invoke(null, new object[] { time })
            ?? throw new InvalidOperationException("TimeOnly conversion returned null.");
    }

    private static bool TryCreateTimeOnly(Type type, string text, out object? value)
    {
        if (TimeSpan.TryParse(text, CultureInfo.InvariantCulture, out var time))
        {
            value = CreateTimeOnly(type, time);
            return true;
        }

        value = null;
        return false;
    }

    private static bool Fail(string text, string expected, out string? error)
    {
        error = $"'{text}' is not {expected}.";
        return false;
    }
}
