using System.ComponentModel;
using System.Diagnostics;
using Nadabas.Setup;

namespace Nadabas.Setup.Tests;

public sealed class ExcelProcessProbeTests
{
    [Fact]
    public void EmptyProcessInventoryMeansNotRunning() =>
        Assert.False(ExcelProcessProbe.HasProcesses(() => []));

    [Fact]
    public void NonEmptyInventoryMeansRunning() =>
        Assert.True(ExcelProcessProbe.HasProcesses(() => [Process.GetCurrentProcess()]));

    [Fact]
    public void DetectionFailurePreventsReplacement()
    {
        var transaction = TestFileTransaction.Create(Path.GetTempPath());
        try
        {
            File.WriteAllText(transaction.Source, "new");
            File.WriteAllText(transaction.Target, "old");
            Assert.Throws<Win32Exception>(() => TestUpdateCoordinator.Apply(transaction, "2", new(),
                () => ExcelProcessProbe.HasProcesses(() => throw new Win32Exception("Access denied"))));
            Assert.Equal("old", File.ReadAllText(transaction.Target));
            Assert.False(File.Exists(transaction.Backup));
            Assert.False(File.Exists(Path.Combine(transaction.DirectoryPath, "transaction.json")));
        }
        finally { Directory.Delete(transaction.DirectoryPath, recursive: true); }
    }

    [Fact]
    public void LiveCommandRejectsRelativePaths() =>
        Assert.Equal(2, Command.Run(["rehearse-update-live", "--source", "relative.xlam", "--output", Path.GetTempPath()], new StringWriter(), new StringWriter()));
}
