namespace DocWeave.Tests;

public sealed class LibraryReferenceTests
{
    [Fact]
    public void DocWeaveAssemblyCanBeLoaded()
    {
        Type markerType = typeof(global::DocWeave.Csv.CsvReader);

        Assert.Equal("DocWeave", markerType.Assembly.GetName().Name);
    }
}
