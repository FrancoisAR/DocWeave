namespace DocWeave.Tests.Excel;

/// <summary>
/// Locates the source workbooks that the import tests read. They are committed files in <c>TestData/Excel</c>, copied next to the
/// test assembly, not workbooks the tests generate. Each one is described in <c>TestData/Excel/README.md</c>.
/// </summary>
internal static class ExcelTestData
{
    public static string PathOf(string fileName)
    {
        return Path.Combine(AppContext.BaseDirectory, "TestData", "Excel", fileName);
    }

    public static byte[] Load(string fileName)
    {
        return File.ReadAllBytes(PathOf(fileName));
    }
}
