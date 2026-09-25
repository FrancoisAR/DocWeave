namespace System.Linq;

internal static class LinqCompatibilityExtensions
{
    public static IEnumerable<TSource> DistinctBy<TSource, TKey>(
        this IEnumerable<TSource> source,
        Func<TSource, TKey> keySelector)
    {
        if (source is null)
        {
            throw new ArgumentNullException(nameof(source));
        }

        if (keySelector is null)
        {
            throw new ArgumentNullException(nameof(keySelector));
        }

        var seenKeys = new HashSet<TKey>();
        foreach (var item in source)
        {
            if (seenKeys.Add(keySelector(item)))
            {
                yield return item;
            }
        }
    }

    public static HashSet<TSource> ToHashSet<TSource>(this IEnumerable<TSource> source)
    {
        return new HashSet<TSource>(source);
    }

    public static HashSet<TSource> ToHashSet<TSource>(
        this IEnumerable<TSource> source,
        IEqualityComparer<TSource> comparer)
    {
        return new HashSet<TSource>(source, comparer);
    }
}
