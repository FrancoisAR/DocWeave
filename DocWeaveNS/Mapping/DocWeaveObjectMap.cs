using System.Linq.Expressions;

namespace DocWeave;

/// <summary>
/// Says how the columns of a file map to the properties of <typeparamref name="T"/>, in code instead of with attributes. Use it
/// for a class you cannot put attributes on, or to map the same class differently for different files.
/// </summary>
/// <typeparam name="T">The type the rows are read into.</typeparam>
/// <remarks>
/// A setting made here replaces the same property's <see cref="DocWeaveColumnAttribute"/> or <see cref="DocWeaveIgnoreAttribute"/>.
/// Properties that are not mentioned keep their attributes, or are matched by name. For a record with a positional constructor,
/// name the property that the constructor parameter creates, for example <c>x =&gt; x.Name</c> for <c>Name</c>.
/// </remarks>
public sealed class DocWeaveObjectMap<T>
{
    internal DocWeaveObjectMap()
    {
    }

    internal ObjectMapSettings Settings { get; } = new();

    /// <summary>
    /// Runs the caller's configuration and returns what it set.
    /// </summary>
    internal static ObjectMapSettings Build(Action<DocWeaveObjectMap<T>> configure)
    {
        NetStandardGuard.ThrowIfNull(configure, nameof(configure));
        var map = new DocWeaveObjectMap<T>();
        configure(map);
        return map.Settings;
    }

    /// <summary>
    /// Fills the property from the column with this header.
    /// </summary>
    /// <param name="property">The property, written as <c>x =&gt; x.Name</c>.</param>
    /// <param name="header">The header text. Compared ignoring case, spaces, underscores and hyphens.</param>
    /// <param name="required">True to make a blank value, and a missing column, an error. False leaves the property's own setting.</param>
    public DocWeaveObjectMap<T> Column<TProperty>(Expression<Func<T, TProperty>> property, string header, bool required = false)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(header, nameof(header));
        var member = Settings.For(NameOf(property));
        member.Ignore = false;
        member.Header = header;
        member.Ordinal = 0;
        ApplyRequired(member, required);
        return this;
    }

    /// <summary>
    /// Fills the property from a column by position.
    /// </summary>
    /// <param name="property">The property, written as <c>x =&gt; x.Name</c>.</param>
    /// <param name="ordinal">
    /// The one-based position among the imported columns (1 is the first). Counted after any column selection, and from the start
    /// cell of an Excel import.
    /// </param>
    /// <param name="required">True to make a blank value, and a missing column, an error. False leaves the property's own setting.</param>
    public DocWeaveObjectMap<T> Column<TProperty>(Expression<Func<T, TProperty>> property, int ordinal, bool required = false)
    {
        NetStandardGuard.ThrowIfLessThan(ordinal, 1, nameof(ordinal));
        var member = Settings.For(NameOf(property));
        member.Ignore = false;
        member.Ordinal = ordinal;
        member.Header = null;
        ApplyRequired(member, required);
        return this;
    }

    /// <summary>
    /// Sets whether a blank value in the property's column, or a missing column, is an error.
    /// </summary>
    /// <param name="property">The property, written as <c>x =&gt; x.Name</c>.</param>
    /// <param name="required">True to require a value. False to allow blanks, even if an attribute says otherwise.</param>
    public DocWeaveObjectMap<T> Required<TProperty>(Expression<Func<T, TProperty>> property, bool required = true)
    {
        Settings.For(NameOf(property)).Required = required;
        return this;
    }

    /// <summary>
    /// Leaves the property out. It is never filled from the file.
    /// </summary>
    /// <param name="property">The property, written as <c>x =&gt; x.Name</c>.</param>
    public DocWeaveObjectMap<T> Ignore<TProperty>(Expression<Func<T, TProperty>> property)
    {
        Settings.For(NameOf(property)).Ignore = true;
        return this;
    }

    private static void ApplyRequired(ObjectMapSettings.Member member, bool required)
    {
        if (required)
        {
            member.Required = true;
        }
    }

    private static string NameOf<TProperty>(Expression<Func<T, TProperty>> property)
    {
        NetStandardGuard.ThrowIfNull(property, nameof(property));
        if (property.Body is MemberExpression { Member: System.Reflection.PropertyInfo info, Expression: ParameterExpression })
        {
            return info.Name;
        }

        throw new ArgumentException("Use a simple property access such as x => x.Name.", nameof(property));
    }
}

/// <summary>
/// The settings collected by a <see cref="DocWeaveObjectMap{T}"/>, by property name.
/// </summary>
internal sealed class ObjectMapSettings
{
    private readonly Dictionary<string, Member> members = new(StringComparer.OrdinalIgnoreCase);

    public sealed class Member
    {
        public string? Header { get; set; }

        public int Ordinal { get; set; }

        public bool? Required { get; set; }

        public bool Ignore { get; set; }
    }

    public Member For(string name)
    {
        if (!members.TryGetValue(name, out var member))
        {
            member = new Member();
            members[name] = member;
        }

        return member;
    }

    public Member? Find(string name)
    {
        return members.TryGetValue(name, out var member) ? member : null;
    }
}
