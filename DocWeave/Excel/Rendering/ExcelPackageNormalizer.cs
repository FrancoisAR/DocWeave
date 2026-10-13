using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocWeave.Excel;

/// <summary>
/// Rewrites part paths and content types after the SDK has saved the package, so pivot caches land where Excel expects them.
/// </summary>
internal static class ExcelPackageNormalizer
{
    internal static MemoryStream NormalizePackage(Stream source)
    {
        var output = new MemoryStream();
        using (var inputArchive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true))
        using (var outputArchive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var inputEntry in inputArchive.Entries)
            {
                // The pivot cache parts arrive under a top-level pivotCache/ folder. Move them under xl/ and rewrite the
                // content types and relationships that point at them (below). Workaround for the layout the saved package has.
                var outputName = inputEntry.FullName.StartsWith("pivotCache/", StringComparison.Ordinal)
                    ? $"xl/{inputEntry.FullName}"
                    : inputEntry.FullName;

                var outputEntry = outputArchive.CreateEntry(outputName, CompressionLevel.Optimal);
                using var inputStream = inputEntry.Open();
                using var outputStream = outputEntry.Open();

                if (ShouldRewriteXmlEntry(inputEntry.FullName))
                {
                    using var reader = new StreamReader(inputStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                    using var writer = new StreamWriter(outputStream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                    writer.Write(RewritePackageXml(inputEntry.FullName, reader.ReadToEnd()));
                }
                else
                {
                    inputStream.CopyTo(outputStream);
                }
            }
        }

        output.Position = 0;
        return output;
    }

    private static bool ShouldRewriteXmlEntry(string entryName)
    {
        // Only the files that name other parts need their text changed: the content type list, and the relationship
        // files that point at the pivot parts that moved.
        return string.Equals(entryName, "[Content_Types].xml", StringComparison.Ordinal)
            || string.Equals(entryName, "xl/_rels/workbook.xml.rels", StringComparison.Ordinal)
            || entryName.StartsWith("xl/worksheets/_rels/", StringComparison.Ordinal)
            || entryName.StartsWith("xl/pivotTables/_rels/", StringComparison.Ordinal)
            || entryName.StartsWith("pivotCache/_rels/", StringComparison.Ordinal);
    }

    private static string RewritePackageXml(string entryName, string text)
    {
        if (string.Equals(entryName, "[Content_Types].xml", StringComparison.Ordinal))
        {
            // Point the content types at the new pivot cache location, and make sure /xl/workbook.xml is declared once,
            // with the workbook content type.
            var rewritten = text
                .Replace("PartName=\"/pivotCache/", "PartName=\"/xl/pivotCache/", StringComparison.Ordinal)
                .Replace(
                    "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" />",
                    "ContentType=\"application/xml\" />",
                    StringComparison.Ordinal);

            if (!rewritten.Contains("PartName=\"/xl/workbook.xml\"", StringComparison.Ordinal))
            {
                rewritten = rewritten.Replace(
                    "</Types>",
                    "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" /></Types>",
                    StringComparison.Ordinal);
            }

            return rewritten;
        }

        if (string.Equals(entryName, "xl/_rels/workbook.xml.rels", StringComparison.Ordinal))
        {
            // The workbook relationships now use a target relative to xl/. The other relationship files below get the
            // equivalent change for the pivot table, worksheet and pivot cache parts.
            return text.Replace("Target=\"/pivotCache/", "Target=\"pivotCache/", StringComparison.Ordinal);
        }

        if (entryName.StartsWith("xl/pivotTables/_rels/", StringComparison.Ordinal))
        {
            return text.Replace("Target=\"/pivotCache/", "Target=\"../pivotCache/", StringComparison.Ordinal);
        }

        if (entryName.StartsWith("xl/worksheets/_rels/", StringComparison.Ordinal))
        {
            return text.Replace("Target=\"/xl/pivotTables/", "Target=\"../pivotTables/", StringComparison.Ordinal);
        }

        if (entryName.StartsWith("pivotCache/_rels/", StringComparison.Ordinal))
        {
            return text.Replace("Target=\"/pivotCache/", "Target=\"", StringComparison.Ordinal);
        }

        return text;
    }
}
