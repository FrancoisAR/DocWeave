namespace DocWeave.Tests;

internal sealed class TestProgress<T> : IProgress<T>
{
    private readonly List<T> reports = [];

    public IReadOnlyList<T> Reports => reports;

    public void Report(T value)
    {
        reports.Add(value);
    }
}
