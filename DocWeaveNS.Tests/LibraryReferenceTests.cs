namespace DocWeaveNS.Tests;

public sealed class LibraryReferenceTests
{
    [Fact]
    public void DocWeaveNSAssemblyCanBeLoaded()
    {
        Type markerType = typeof(global::DocWeave.Csv.CsvReader);

        Assert.Equal("DocWeaveNS", markerType.Assembly.GetName().Name);
    }
}
