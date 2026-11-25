using System.Text.Json;

namespace Nadabas.Setup;

internal static class ReleaseBaseline
{
    internal static string RequireAddin(string source)
    {
        TestFileTransaction.CheckPath(source);
        using var stream = typeof(ReleaseBaseline).Assembly.GetManifestResourceStream("Nadabas.ReleaseBaseline.json")!;
        using var baseline = JsonDocument.Parse(stream);
        var expected = baseline.RootElement.GetProperty("files").EnumerateArray()
            .Single(f => f.GetProperty("role").GetString() == "excel-addin").GetProperty("sha256").GetString()!;
        if (TestFileTransaction.Hash(source) != expected)
            throw new IOException("Source does not match the pinned 6.02.005 add-in.");
        return expected;
    }
}
