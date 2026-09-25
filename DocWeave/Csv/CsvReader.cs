using System.Text;

namespace DocWeave.Csv;

/// <summary>
/// Entry point for reading CSV from a file, a stream, bytes or text. Each method returns a <see cref="CsvReadBuilder"/> to set options and read.
/// </summary>
public static class CsvReader
{
    /// <summary>
    /// Starts a CSV import from a file. The file is opened when reading starts and closed when it ends.
    /// </summary>
    public static CsvReadBuilder FromFile(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return new CsvReadBuilder(() => File.OpenRead(path), ownsStream: true, Path.GetFileName(path));
    }

    /// <summary>
    /// Starts a CSV import from a caller-owned stream. The stream is read from its current position
    /// and is left open, so the caller decides when to dispose it. It can be read only once.
    /// </summary>
    public static CsvReadBuilder FromStream(Stream stream, string? label = null)
    {
        ArgumentNullException.ThrowIfNull(stream);
        return new CsvReadBuilder(() => stream, ownsStream: false, label);
    }

    /// <summary>
    /// Starts a CSV import from bytes held in memory.
    /// </summary>
    /// <param name="bytes">The file content.</param>
    /// <param name="label">A name for the source, used in messages.</param>
    public static CsvReadBuilder FromBytes(byte[] bytes, string? label = null)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        return new CsvReadBuilder(() => new MemoryStream(bytes), ownsStream: true, label);
    }

    /// <summary>
    /// Starts a CSV import from text. The text is treated as UTF-8.
    /// </summary>
    /// <param name="text">The file content.</param>
    /// <param name="label">A name for the source, used in messages.</param>
    public static CsvReadBuilder FromText(string text, string? label = null)
    {
        ArgumentNullException.ThrowIfNull(text);
        return FromBytes(Encoding.UTF8.GetBytes(text), label);
    }
}
