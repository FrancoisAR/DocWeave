namespace DocWeave;

/// <summary>
/// Says which imported column fills a property, or a constructor parameter, when reading rows into objects with <c>ReadObjects</c>
/// or <c>EnumerateObjects</c>. On a positional record, write it on the parameter (<c>[DocWeaveColumn("Sales Rep")] string Rep</c>) or on the
/// property (<c>[property: DocWeaveColumn("Sales Rep")]</c>).
/// Without this attribute a property or parameter is matched to the column whose header has the same name.
/// </summary>
/// <remarks>
/// A column is found by, in order: <see cref="Ordinal"/> if it is set, then <see cref="Name"/>, then the property name.
/// Header names are compared ignoring case, spaces, underscores and hyphens, so a property called <c>InvoiceNo</c> matches a
/// header <c>Invoice No</c> without any attribute.
/// </remarks>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Parameter, AllowMultiple = false, Inherited = true)]
public sealed class DocWeaveColumnAttribute : Attribute
{
    /// <summary>
    /// Matches the property to a column by header name, or by property name when no name is given.
    /// </summary>
    public DocWeaveColumnAttribute()
    {
    }

    /// <summary>
    /// Matches the property to the column with this header.
    /// </summary>
    /// <param name="name">The header text.</param>
    public DocWeaveColumnAttribute(string name)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(name, nameof(name));
        Name = name;
    }

    /// <summary>
    /// The header text to match. When null the property name is used.
    /// </summary>
    public string? Name { get; set; }

    /// <summary>
    /// The one-based position of the column among the imported columns (1 is the first). Zero, the default, means "not set".
    /// The position is counted after any column selection, and from the start cell of an Excel import, so for a sheet imported
    /// from <c>B4</c>, ordinal 1 is column B.
    /// </summary>
    public int Ordinal { get; set; }

    /// <summary>
    /// When true, a blank value in this column is an error, and so is a file that has no such column. Blank values are otherwise
    /// left as the property's default.
    /// </summary>
    public bool Required { get; set; }
}

/// <summary>
/// Leaves a property, or a constructor parameter, out when reading rows into objects. A left-out constructor parameter receives its
/// default value.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Parameter, AllowMultiple = false, Inherited = true)]
public sealed class DocWeaveIgnoreAttribute : Attribute
{
}
