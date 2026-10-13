namespace DocWeave.Tests.Excel;

internal static class ExcelTestOutput
{
    public static void SaveWorkbook(string testName, byte[] bytes)
    {
        var outputFolder = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "TestOutput",
            "Excel"));

        Directory.CreateDirectory(outputFolder);
        var outputPath = Path.Combine(outputFolder, $"{testName}.xlsx");
        try
        {
            File.WriteAllBytes(outputPath, bytes);
        }
        catch (IOException)
        {
            // Test output files are often opened manually while developing the exporter.
            // Keep the test run useful by writing a clearly named fallback copy.
            File.WriteAllBytes(Path.Combine(outputFolder, $"{testName}.latest.xlsx"), bytes);
        }
    }
}
