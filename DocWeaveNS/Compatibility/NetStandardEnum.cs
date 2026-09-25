namespace DocWeave;

internal static class NetStandardEnum
{
    public static TEnum Parse<TEnum>(string value, bool ignoreCase) where TEnum : struct
    {
        return (TEnum)Enum.Parse(typeof(TEnum), value, ignoreCase);
    }
}
