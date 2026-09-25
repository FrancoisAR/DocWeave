using System.Text;

namespace DocWeave.Excel;

/// <summary>
/// Entry point for reading Excel workbooks from common server-side sources.
/// </summary>
public static class ExcelReader
{
    /// <summary>
    /// Starts an Excel import from a workbook on disk.
    /// </summary>
    public static ExcelReadBuilder FromFile(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return new ExcelReadBuilder(() => File.OpenRead(path), Path.GetFileName(path));
    }

    /// <summary>
    /// Starts an Excel import from an existing stream.
    /// </summary>
    public static ExcelReadBuilder FromStream(Stream stream, string? label = null)
    {
        ArgumentNullException.ThrowIfNull(stream);
        return new ExcelReadBuilder(() => stream, label);
    }

    /// <summary>
    /// Starts an Excel import from binary workbook content, such as a Web API upload body.
    /// </summary>
    public static ExcelReadBuilder FromBytes(byte[] bytes, string? label = null)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        return new ExcelReadBuilder(() => new MemoryStream(bytes), label);
    }

    /// <summary>
    /// Starts an Excel import from a base64 encoded workbook payload.
    /// </summary>
    public static ExcelReadBuilder FromBase64(string base64, string? label = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(base64);
        return FromBytes(Convert.FromBase64String(base64), label);
    }

    /// <summary>
    /// Starts an Excel import from text that represents a base64 encoded workbook payload.
    /// </summary>
    public static ExcelReadBuilder FromBase64Text(string base64Text, Encoding? encoding = null, string? label = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(base64Text);
        var text = (encoding ?? Encoding.UTF8).GetString((encoding ?? Encoding.UTF8).GetBytes(base64Text));
        return FromBase64(text, label);
    }
}
