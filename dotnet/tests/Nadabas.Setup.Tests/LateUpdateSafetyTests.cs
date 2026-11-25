using Nadabas.Setup;

namespace Nadabas.Setup.Tests;

public sealed class LateUpdateSafetyTests : IDisposable
{
    private readonly TestFileTransaction transaction = TestFileTransaction.Create(Path.GetTempPath());

    public LateUpdateSafetyTests()
    {
        File.WriteAllText(transaction.Source, "new");
        File.WriteAllText(transaction.Target, "old");
    }

    [Fact]
    public void LockedNewFileIsPreservedForRecoveryAfterReaderCloses()
    {
        if (!OperatingSystem.IsWindows()) return;
        FileStream? reader = null;
        TransactionResult result;
        try
        {
            result = TestUpdateCoordinator.Apply(transaction, "2", new(), () => false, phase =>
            {
                if (phase != "replaced") return;
                reader = new FileStream(transaction.Target, FileMode.Open, FileAccess.Read, FileShare.Read);
                throw new IOException("Failure while another process holds the new file");
            });
        }
        finally { reader?.Dispose(); }
        Assert.Equal("recovery-required", result.Status);
        Assert.Equal("new", File.ReadAllText(transaction.Target));
        Assert.Equal("old", File.ReadAllText(transaction.Backup));
        Assert.Equal("rolled-back", transaction.Recover().Status);
        Assert.Equal("old", File.ReadAllText(transaction.Target));
    }

    [Fact]
    public void ReaderArrivingAfterPreparationPreventsReplacement()
    {
        if (!OperatingSystem.IsWindows()) return;
        FileStream? reader = null;
        TransactionResult result;
        var probes = 0;
        try
        {
            result = TestUpdateCoordinator.Apply(transaction, "2", new(), () =>
            {
                if (++probes == 3) reader = new FileStream(transaction.Target, FileMode.Open, FileAccess.Read, FileShare.Read);
                return false;
            });
        }
        finally { reader?.Dispose(); }
        Assert.Equal("rolled-back", result.Status);
        Assert.NotNull(result.Error);
        Assert.Equal("old", File.ReadAllText(transaction.Target));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LateExcelStartOrDetectionErrorPreservesOldFile(bool detectionError)
    {
        var probes = 0;
        var result = TestUpdateCoordinator.Apply(transaction, "2", new(), () =>
        {
            if (++probes < 3) return false;
            if (detectionError) throw new InvalidOperationException("Detection failed");
            return true;
        });
        Assert.Equal("rolled-back", result.Status);
        Assert.Contains(detectionError ? "Could not confirm" : "Excel started", result.Error!);
        Assert.Equal("old", File.ReadAllText(transaction.Target));
    }

    [Fact]
    public void FailureAfterRenameLeavesNoTransactionHandlesOpen()
    {
        if (!OperatingSystem.IsWindows()) return;
        var result = TestUpdateCoordinator.Apply(transaction, "2", new(), () => false, phase =>
        {
            if (phase == "replaced") throw new IOException("Injected failure");
        });
        Assert.Equal("rolled-back", result.Status);
        Assert.Equal("old", File.ReadAllText(transaction.Target));
        using var exclusive = new FileStream(transaction.Target, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
    }

    public void Dispose() => Directory.Delete(transaction.DirectoryPath, recursive: true);
}
