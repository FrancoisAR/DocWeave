using System.Reflection;

namespace DocWeave;

/// <summary>
/// Fills objects of type <typeparamref name="T"/> from imported rows. The columns are matched to properties (and, for a type
/// with no parameterless constructor, to constructor parameters) once, from the headers, and each row is then converted with
/// the same plan.
/// </summary>
internal sealed class RowObjectMapper<T>
{
    /// <summary>
    /// One value to fill: a constructor parameter (<c>ParameterIndex</c> 0 or more) or a property (<c>ParameterIndex</c> -1).
    /// </summary>
    private sealed record Slot(Type Type, string Header, bool Required, PropertyInfo? Property, int ParameterIndex, bool HasDefault, string Name);

    private readonly List<Slot> slots = [];
    private readonly List<DocWeaveImportError> columnErrors = [];
    private readonly bool excelSerialDates;
    private ConstructorInfo? constructor;
    private object?[] parameterDefaults = [];

    private RowObjectMapper(bool excelSerialDates)
    {
        this.excelSerialDates = excelSerialDates;
    }

    /// <summary>
    /// Problems with the columns as a whole, found before any row is read. When there are any, no row can be mapped.
    /// </summary>
    public IReadOnlyList<DocWeaveImportError> ColumnErrors => columnErrors;

    /// <summary>
    /// Matches the members of <typeparamref name="T"/> to <paramref name="headers"/>.
    /// </summary>
    /// <param name="headers">The imported column headers, after any column selection.</param>
    /// <param name="excelSerialDates">True for an Excel source, where a number may stand for a date.</param>
    /// <param name="settings">Fluent settings that replace the attributes, or null.</param>
    /// <exception cref="NotSupportedException">
    /// The type cannot be built (no usable public constructor), or a property or parameter has a type that a cell cannot be converted to.
    /// </exception>
    public static RowObjectMapper<T> Create(IReadOnlyList<string> headers, bool excelSerialDates, ObjectMapSettings? settings = null)
    {
        var mapper = new RowObjectMapper<T>(excelSerialDates);
        var type = typeof(T);
        settings ??= new ObjectMapSettings();

        var exact = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var loose = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var header in headers)
        {
            exact.TryAdd(header, header);
            loose.TryAdd(Loosen(header), header);
        }

        var allProperties = type.GetProperties(BindingFlags.Instance | BindingFlags.Public);

        mapper.constructor = ChooseConstructor(type);
        var parameters = mapper.constructor?.GetParameters() ?? [];
        mapper.parameterDefaults = parameters
            .Select(parameter => parameter.HasDefaultValue ? parameter.DefaultValue : DefaultOf(parameter.ParameterType))
            .ToArray();

        // Constructor parameters first. A positional record's parameter and its property share a name, so the property's
        // attribute (written as [property: ...]) counts as well as one on the parameter itself.
        var coveredByConstructor = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < parameters.Length; index++)
        {
            var parameter = parameters[index];
            var name = parameter.Name!;
            coveredByConstructor.Add(name);
            var property = allProperties.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));

            var attribute = parameter.GetCustomAttribute<DocWeaveColumnAttribute>() ?? property?.GetCustomAttribute<DocWeaveColumnAttribute>(inherit: true);
            var ignored = parameter.IsDefined(typeof(DocWeaveIgnoreAttribute)) || (property?.IsDefined(typeof(DocWeaveIgnoreAttribute), inherit: true) ?? false);
            mapper.AddSlot(
                property?.Name ?? name, name, parameter.ParameterType, attribute, ignored, settings, headers, exact, loose,
                property: null, parameterIndex: index, hasDefault: parameter.HasDefaultValue);
        }

        // Then the public settable properties the constructor does not already fill.
        foreach (var property in allProperties)
        {
            if (property.SetMethod is not { IsPublic: true }
                || property.GetIndexParameters().Length > 0
                || coveredByConstructor.Contains(property.Name))
            {
                continue;
            }

            var attribute = property.GetCustomAttribute<DocWeaveColumnAttribute>(inherit: true);
            var ignored = property.IsDefined(typeof(DocWeaveIgnoreAttribute), inherit: true);
            mapper.AddSlot(
                property.Name, property.Name, property.PropertyType, attribute, ignored, settings, headers, exact, loose,
                property: property, parameterIndex: -1, hasDefault: false);
        }

        return mapper;
    }

    /// <summary>
    /// Builds one object from a row. Every problem in the row is reported, not just the first.
    /// </summary>
    /// <returns>True when the row had no problems and an object was built. When false, <paramref name="item"/> is not set.</returns>
    public bool TryMap(int rowNumber, IReadOnlyDictionary<string, string?> values, Action<DocWeaveImportError>? report, out T item)
    {
        item = default!;
        var ok = true;
        var arguments = (object?[])parameterDefaults.Clone();
        var assignments = new List<(PropertyInfo Property, object? Value)>();

        foreach (var slot in slots)
        {
            values.TryGetValue(slot.Header, out var text);

            // Text is kept as it is, spaces included. For anything else, spaces alone count as blank.
            var isBlank = slot.Type == typeof(string)
                ? string.IsNullOrEmpty(text)
                : string.IsNullOrWhiteSpace(text);

            if (isBlank)
            {
                if (slot.Required || !(ValueConversion.AllowsBlank(slot.Type) || slot.HasDefault))
                {
                    report?.Invoke(new DocWeaveImportError(rowNumber, slot.Header, slot.Name, text, "A value is required."));
                    ok = false;
                }

                // A blank leaves the value as the parameter default, or as the constructor or initializer set it.
                continue;
            }

            if (!ValueConversion.TryConvert(text!, slot.Type, excelSerialDates, out var converted, out var message))
            {
                report?.Invoke(new DocWeaveImportError(rowNumber, slot.Header, slot.Name, text, message!));
                ok = false;
                continue;
            }

            if (slot.ParameterIndex >= 0)
            {
                arguments[slot.ParameterIndex] = converted;
            }
            else
            {
                assignments.Add((slot.Property!, converted));
            }
        }

        if (!ok)
        {
            return false;
        }

        try
        {
            // Filled as an object so a struct can be changed in place before it is unboxed.
            var instance = constructor is null ? Activator.CreateInstance(typeof(T))! : constructor.Invoke(arguments);
            foreach (var (property, value) in assignments)
            {
                property.SetValue(instance, value);
            }

            item = (T)instance;
            return true;
        }
        catch (TargetInvocationException exception)
        {
            // The type's own code (a constructor or a setter) refused the values. That is a problem with this row, not a crash.
            report?.Invoke(new DocWeaveImportError(rowNumber, null, null, null, $"The row was rejected while building {typeof(T).Name}: {exception.InnerException?.Message ?? exception.Message}"));
            return false;
        }
    }

    private void AddSlot(
        string memberName,
        string parameterOrPropertyName,
        Type memberType,
        DocWeaveColumnAttribute? attribute,
        bool ignoredByAttribute,
        ObjectMapSettings settings,
        IReadOnlyList<string> headers,
        Dictionary<string, string> exact,
        Dictionary<string, string> loose,
        PropertyInfo? property,
        int parameterIndex,
        bool hasDefault)
    {
        var fluent = settings.Find(memberName);

        // A fluent Column(...) brings back a property an attribute ignored. A fluent Ignore(...) wins over everything.
        var ignored = fluent is not null
            ? fluent.Ignore || (ignoredByAttribute && fluent.Header is null && fluent.Ordinal == 0)
            : ignoredByAttribute;
        if (ignored)
        {
            return;
        }

        if (!ValueConversion.IsSupported(memberType))
        {
            var kind = parameterIndex >= 0 ? "Constructor parameter" : "Property";
            throw new NotSupportedException(
                $"{kind} '{parameterOrPropertyName}' of {typeof(T).Name} has type {memberType.Name}, which cannot be filled from a cell. Mark it with [DocWeaveIgnore] or use a supported type.");
        }

        // Precedence: fluent position, fluent header, attribute position, attribute header, then the member's own name.
        string? header;
        string wanted;
        var ordinal = fluent is { Ordinal: > 0 } ? fluent.Ordinal
            : fluent?.Header is not null ? 0
            : attribute?.Ordinal ?? 0;
        if (ordinal > 0)
        {
            header = ordinal <= headers.Count ? headers[ordinal - 1] : null;
            wanted = $"column {ordinal}";
        }
        else
        {
            wanted = fluent?.Header ?? attribute?.Name ?? parameterOrPropertyName;
            header = exact.TryGetValue(wanted, out var found) ? found
                : loose.TryGetValue(Loosen(wanted), out var looseMatch) ? looseMatch
                : null;
        }

        var required = fluent?.Required ?? attribute?.Required ?? false;
        if (header is null)
        {
            if (required)
            {
                columnErrors.Add(new DocWeaveImportError(0, wanted, parameterOrPropertyName, null, $"Required column '{wanted}' was not found."));
            }

            return;
        }

        slots.Add(new Slot(memberType, header, required, property, parameterIndex, hasDefault, parameterOrPropertyName));
    }

    /// <summary>
    /// A parameterless public constructor is used when there is one (and always for a struct). Otherwise the public constructor
    /// with the most parameters, which for a positional record is the primary constructor.
    /// </summary>
    private static ConstructorInfo? ChooseConstructor(Type type)
    {
        if (type.IsValueType || type.GetConstructor(Type.EmptyTypes) is not null)
        {
            return null;
        }

        if (type.IsAbstract || type.IsInterface)
        {
            throw new NotSupportedException($"{type.Name} is abstract, so rows cannot be read into it. Use a concrete type.");
        }

        var constructors = type.GetConstructors();
        if (constructors.Length == 0)
        {
            throw new NotSupportedException($"{type.Name} has no public constructor, so rows cannot be read into it.");
        }

        var most = constructors.Max(candidate => candidate.GetParameters().Length);
        var best = constructors.Where(candidate => candidate.GetParameters().Length == most).ToList();
        if (best.Count > 1)
        {
            throw new NotSupportedException(
                $"{type.Name} has more than one public constructor with {most} parameters, so it is unclear which to use. Give it a parameterless constructor, or a single main constructor.");
        }

        return best[0];
    }

    private static object? DefaultOf(Type type)
    {
        return type.IsValueType ? Activator.CreateInstance(type) : null;
    }

    /// <summary>
    /// "Invoice No", "invoice_no" and "InvoiceNo" all become "invoiceno".
    /// </summary>
    private static string Loosen(string name)
    {
        return new string(name.Where(character => character is not (' ' or '_' or '-')).ToArray()).ToLowerInvariant();
    }
}
