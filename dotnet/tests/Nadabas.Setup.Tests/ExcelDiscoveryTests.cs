using Nadabas.Setup;

namespace Nadabas.Setup.Tests;

public sealed class ExcelDiscoveryTests : IDisposable
{
    private readonly string root = Path.Combine(Path.GetTempPath(), "nadabas-pe-" + Guid.NewGuid());

    private string Image(string folder, ushort machine = 0x14c, ushort magic = 0x10b, bool dll = false)
    {
        var directory = Directory.CreateDirectory(Path.Combine(root, folder));
        var path = Path.Combine(directory.FullName, "EXCEL.EXE");
        using var stream = File.Create(path);
        using var writer = new BinaryWriter(stream);
        stream.SetLength(1024);
        writer.Write((ushort)0x5a4d);
        stream.Position = 0x3c;
        writer.Write(0x80);
        stream.Position = 0x80;
        writer.Write(0x00004550);
        writer.Write(machine);
        writer.Write((ushort)1);
        stream.Position = 0x94;
        writer.Write((ushort)(magic == 0x20b ? 240 : 224));
        writer.Write((ushort)(dll ? 0x2002 : 0x0002));
        writer.Write(magic);
        return path;
    }

    [Theory]
    [InlineData(0x14c, 0x10b, "x86")]
    [InlineData(0x8664, 0x20b, "x64")]
    public void IdentifiesExecutableNotInspectorArchitecture(int machine, int magic, string expected)
    {
        var path = Image("office", (ushort)machine, (ushort)magic);
        var before = File.ReadAllBytes(path);
        var result = ExcelDiscovery.Inspect([new("HKLM32", path)]);
        Assert.Equal("verified", result.Status);
        Assert.Equal(expected, result.Selected!.Architecture);
        Assert.Equal(before, File.ReadAllBytes(path));
    }

    [Fact]
    public void SharedRegistryPathsAreDeduplicatedWithEvidence()
    {
        var path = Image("office");
        var result = ExcelDiscovery.Inspect([new("HKLM32", path), new("HKLM64", '"' + path + '"')]);
        Assert.Equal("verified", result.Status);
        Assert.Equal(2, Assert.Single(result.Candidates).Sources.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CompetingRegistrationNeverSelectsSilently(bool missing)
    {
        var first = Image("one");
        var second = missing ? Path.Combine(root, "missing", "EXCEL.EXE") : Image("two", 0x8664, 0x20b);
        var result = ExcelDiscovery.Inspect([new("user", first), new("machine", second)]);
        Assert.Equal("ambiguous", result.Status);
        Assert.Null(result.Selected);
    }

    [Theory]
    [InlineData(0xaa64, 0x20b, false, "unsupported-image")]
    [InlineData(0x8664, 0x10b, false, "unsupported-image")]
    [InlineData(0x14c, 0x10b, true, "invalid-image")]
    public void RejectsUnsupportedOrInconsistentImages(int machine, int magic, bool dll, string status)
    {
        var result = ExcelDiscovery.Inspect([new("test", Image("office", (ushort)machine, (ushort)magic, dll))]);
        Assert.Equal(status, result.Status);
        Assert.Null(result.Selected);
    }

    [Fact]
    public void MissingExecutableDoesNotCreateDirectories()
    {
        var result = ExcelDiscovery.Inspect([new("test", Path.Combine(root, "EXCEL.EXE"))]);
        Assert.Equal("missing", result.Status);
        Assert.Null(result.Selected);
        Assert.False(Directory.Exists(root));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(100)]
    [InlineData(160)]
    public void TruncatedHeaderCannotSelectAnInstallation(int length)
    {
        var path = Image("office");
        using (var stream = File.OpenWrite(path)) stream.SetLength(length);
        var result = ExcelDiscovery.Inspect([new("test", path)]);
        Assert.NotEqual("verified", result.Status);
        Assert.Null(result.Selected);
    }

    [Theory]
    [InlineData("EXCEL.EXE")]
    [InlineData("relative/EXCEL.EXE")]
    public void RelativePathsAreRejected(string path) =>
        Assert.Equal("invalid-path", ExcelDiscovery.Inspect([new("test", path)]).Status);

    [Fact]
    public void CommandArgumentsAreNotExecutedOrInterpreted()
    {
        var path = Image("office");
        Assert.Equal("invalid-path", ExcelDiscovery.Inspect([new("test", '"' + path + "\" /automation")]).Status);
    }

    [Fact]
    public void NoRegistrationIsExplicit() =>
        Assert.Equal("not-found", ExcelDiscovery.Inspect([new("test", null)]).Status);

    public void Dispose()
    {
        if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
    }
}
