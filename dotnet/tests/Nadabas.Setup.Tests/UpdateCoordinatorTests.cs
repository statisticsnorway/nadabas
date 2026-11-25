using Nadabas.Setup;

namespace Nadabas.Setup.Tests;

public sealed class UpdateCoordinatorTests : IDisposable
{
    private readonly TestFileTransaction transaction = TestFileTransaction.Create(Path.GetTempPath());

    public UpdateCoordinatorTests()
    {
        File.WriteAllText(transaction.Source, "incoming package");
        File.WriteAllText(transaction.Target, "previous package");
    }

    [Fact]
    public void DeferredUpdateCanResumeWithoutLeavingATransactionBehind()
    {
        Assert.Equal("deferred-excel-running", TestUpdateCoordinator.Apply(transaction, "2", new(), () => true).Status);
        Assert.Equal("previous package", File.ReadAllText(transaction.Target));
        Assert.False(File.Exists(transaction.Backup));
        Assert.Equal("installed", TestUpdateCoordinator.Apply(transaction, "2", new(), () => false).Status);
        Assert.Equal("previous package", File.ReadAllText(transaction.Backup));
        Assert.Equal("incoming package", File.ReadAllText(transaction.Target));
    }

    [Fact]
    public void ExcelStartingAfterPlanningDefersBeforeReplacement()
    {
        var probes = 0;
        Assert.Equal("deferred-excel-running", TestUpdateCoordinator.Apply(transaction, "2", new(), () => ++probes == 2).Status);
        Assert.Equal("previous package", File.ReadAllText(transaction.Target));
        Assert.False(File.Exists(Path.Combine(transaction.DirectoryPath, "transaction.json")));
    }

    [Theory]
    [InlineData(true, false, null, "disabled-by-policy")]
    [InlineData(false, true, null, "installed")]
    [InlineData(true, true, "1", "pinned-version-mismatch")]
    [InlineData(true, false, "2", "disabled-by-policy")]
    [InlineData(true, true, "2", "installed")]
    public void ItPolicyTakesPrecedence(bool userEnabled, bool itEnabled, string? pinned, string expected)
    {
        var result = TestUpdateCoordinator.Apply(transaction, "2", new(userEnabled, itEnabled, pinned), () => false);
        Assert.Equal(expected, result.Status);
        Assert.Equal(expected == "installed" ? "incoming package" : "previous package", File.ReadAllText(transaction.Target));
        Assert.Equal(expected == "installed", File.Exists(transaction.Backup));
    }

    [Fact]
    public void FailedUpgradeRestoresPreviousBytesAndCanBeRecoveredAgain()
    {
        var result = TestUpdateCoordinator.Apply(transaction, "2", new(), () => false,
            phase => { if (phase == "replaced") throw new IOException("Injected failure"); });
        Assert.Equal("rolled-back", result.Status);
        Assert.Equal("previous package", File.ReadAllText(transaction.Target));
        Assert.Equal("rolled-back", new TestFileTransaction(transaction.DirectoryPath).Recover().Status);
    }

    [Fact]
    public void InvalidCommandDoesNotCreateOutput()
    {
        var output = Path.Combine(transaction.DirectoryPath, "not-created");
        Assert.Equal(2, Command.Run(["rehearse-update", "--source", "relative.xlam", "--output", output], new StringWriter(), new StringWriter()));
        Assert.False(Directory.Exists(output));
    }

    public void Dispose() => Directory.Delete(transaction.DirectoryPath, recursive: true);
}
