using System.Text.Json;
using Nadabas.Setup;

namespace Nadabas.Setup.Tests;

public sealed class DiagnosticsTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "nadabas-diagnostics-" + Guid.NewGuid());

    public DiagnosticsTests() => Directory.CreateDirectory(directory);

    [Theory]
    [InlineData("NADABAS.xlam", true)]
    [InlineData("nadabas.6.02.004.XLAM", true)]
    [InlineData("NADABAS-backup.xlam", false)]
    [InlineData("NADABAS.xlsm", false)]
    [InlineData("other.xlam", false)]
    public void MatchesLegacyInstallerFilenameConvention(string name, bool expected) =>
        Assert.Equal(expected, Diagnostics.IsNadabasAddin(name));

    [Fact]
    public void InventoryDoesNotChangeFilesOrSearchSubdirectories()
    {
        var path = Path.Combine(directory, "NADABAS.xlam");
        byte[] bytes = [0, 1, 2, 255];
        File.WriteAllBytes(path, bytes);
        File.WriteAllText(Path.Combine(directory, "other.xlam"), "unrelated");
        var nested = Directory.CreateDirectory(Path.Combine(directory, "nested"));
        File.WriteAllText(Path.Combine(nested.FullName, "NADABAS.old.xlam"), "nested");
        var before = Directory.GetFiles(directory, "*", SearchOption.AllDirectories).Order().ToArray();

        var findings = Diagnostics.InspectAddins(directory);

        Assert.Equal(path, Assert.Single(findings, f => f.Id == "addin.file").Value);
        Assert.Equal(bytes, File.ReadAllBytes(path));
        Assert.Equal(before, Directory.GetFiles(directory, "*", SearchOption.AllDirectories).Order().ToArray());
    }

    [Fact]
    public void MultipleCandidatesAreReportedWithoutSelectingOrDeletingOne()
    {
        File.WriteAllText(Path.Combine(directory, "NADABAS.xlam"), "first");
        File.WriteAllText(Path.Combine(directory, "NADABAS.6.01.xlam"), "second");
        var findings = Diagnostics.InspectAddins(directory);
        Assert.Equal(2, findings.Count(f => f.Id == "addin.file"));
        Assert.Contains(findings, f => f.Id == "addin.multiple" && f.Status == "warning");
        Assert.Equal(2, Directory.GetFiles(directory).Length);
    }

    [Fact]
    public void MissingDirectoryIsNotCreated()
    {
        var missing = Path.Combine(directory, "missing");
        Assert.Contains(Diagnostics.InspectAddins(missing), f => f.Status == "not-found");
        Assert.False(Directory.Exists(missing));
    }

    [Fact]
    public void FileUsedAsDirectoryIsAnInspectionError()
    {
        var path = Path.Combine(directory, "file");
        File.WriteAllText(path, "not a directory");
        Assert.Contains(Diagnostics.InspectAddins(path), f => f.Status == "error");
    }

    public static TheoryData<string[]> InvalidArguments => new()
    {
        Array.Empty<string>(),
        new[] { "install" },
        new[] { "diagnose", "--addin-directory" },
        new[] { "diagnose", "--addin-directory", "relative" },
        new[] { "diagnose", "--unknown", "value" }
    };

    [Theory]
    [MemberData(nameof(InvalidArguments))]
    public void InvalidArgumentsHaveNoReport(string[] args)
    {
        var output = new StringWriter();
        var error = new StringWriter();
        Assert.Equal(2, Command.Run(args, output, error));
        Assert.Empty(output.ToString());
        Assert.NotEmpty(error.ToString());
    }

    [Fact]
    public void CommandProducesVersionedJsonAndHonestCompletionStatus()
    {
        var output = new StringWriter();
        var error = new StringWriter();
        var exitCode = Command.Run(["diagnose", "--addin-directory", directory], output, error);
        using var json = JsonDocument.Parse(output.ToString());
        Assert.Equal(1, json.RootElement.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(json.RootElement.GetProperty("complete").GetBoolean() ? 0 : 1, exitCode);
        Assert.Contains(json.RootElement.GetProperty("findings").EnumerateArray(),
            f => f.GetProperty("id").GetString() == "excel.runtime" && f.GetProperty("status").GetString() == "not-tested");
        Assert.Empty(error.ToString());
    }

    public void Dispose() => Directory.Delete(directory, recursive: true);
}
