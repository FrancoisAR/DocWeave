namespace DocWeave.Tests.Csv;

internal static class CsvTestOutput
{
    public static void SaveCsv(string testName, string csv)
    {
        var outputFolder = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "TestOutput",
            "Csv"));

        Directory.CreateDirectory(outputFolder);
        File.WriteAllText(Path.Combine(outputFolder, $"{testName}.csv"), csv);
    }
}
