using Nadabas.Setup;

namespace Nadabas.Setup.Tests;

public sealed class TestFileTransactionTests : IDisposable
{
    private readonly TestFileTransaction transaction = TestFileTransaction.Create(Path.GetTempPath());

    public TestFileTransactionTests() => File.WriteAllText(transaction.Source, "new bytes");

    [Fact]
    public void ReplacementAndExplicitRecoveryPreserveExactBytes()
    {
        File.WriteAllText(transaction.Target, "old bytes");
        var plan = transaction.Plan();
        Assert.Equal("installed", transaction.Apply(plan).Status);
        Assert.Equal("new bytes", File.ReadAllText(transaction.Target));
        Assert.Equal("old bytes", File.ReadAllText(transaction.Backup));
        Assert.Equal("rolled-back", new TestFileTransaction(transaction.DirectoryPath).Recover().Status);
        Assert.Equal("old bytes", File.ReadAllText(transaction.Target));
        Assert.Equal("rolled-back", transaction.Recover().Status);
    }

    [Fact]
    public void FreshInstallRecoveryRestoresAbsenceWithoutDeletingNewBytes()
    {
        Assert.Equal("installed", transaction.Apply(transaction.Plan()).Status);
        Assert.Equal("rolled-back", transaction.Recover().Status);
        Assert.False(File.Exists(transaction.Target));
        Assert.Equal("new bytes", File.ReadAllText(Path.Combine(transaction.DirectoryPath, "rolled-back-new.xlam")));
    }

    [Fact]
    public void IdenticalFileHasNoBackupOrReplacement()
    {
        File.Copy(transaction.Source, transaction.Target);
        Assert.Equal("unchanged", transaction.Apply(transaction.Plan()).Status);
        Assert.False(File.Exists(transaction.Backup));
        Assert.False(File.Exists(Path.Combine(transaction.DirectoryPath, "transaction.json")));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void StalePlanIsRejected(bool changeSource)
    {
        File.WriteAllText(transaction.Target, "old bytes");
        var plan = transaction.Plan();
        File.WriteAllText(changeSource ? transaction.Source : transaction.Target, "changed later");
        Assert.Equal("blocked", transaction.Apply(plan).Status);
        Assert.Equal(changeSource ? "old bytes" : "changed later", File.ReadAllText(transaction.Target));
        Assert.False(File.Exists(transaction.Backup));
    }

    [Fact]
    public void LockedTargetIsNotOverwritten()
    {
        File.WriteAllText(transaction.Target, "old bytes");
        var plan = transaction.Plan();
        using (var locked = new FileStream(transaction.Target, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            Assert.Equal("blocked", transaction.Apply(plan).Status);
        Assert.Equal("old bytes", File.ReadAllText(transaction.Target));
    }

    [Fact]
    public void CorruptExistingBackupIsNeverOverwritten()
    {
        File.WriteAllText(transaction.Target, "old bytes");
        File.WriteAllText(transaction.Backup, "unrelated backup");
        Assert.Equal("blocked", transaction.Apply(transaction.Plan()).Status);
        Assert.Equal("old bytes", File.ReadAllText(transaction.Target));
        Assert.Equal("unrelated backup", File.ReadAllText(transaction.Backup));
    }

    [Fact]
    public void ReadOnlyTargetIsPreserved()
    {
        File.WriteAllText(transaction.Target, "old bytes");
        var plan = transaction.Plan();
        File.SetAttributes(transaction.Target, FileAttributes.ReadOnly);
        try
        {
            Assert.Equal("blocked", transaction.Apply(plan).Status);
            Assert.Equal("old bytes", File.ReadAllText(transaction.Target));
            Assert.True((File.GetAttributes(transaction.Target) & FileAttributes.ReadOnly) != 0);
        }
        finally { File.SetAttributes(transaction.Target, FileAttributes.Normal); }
    }

    [Fact]
    public void RecoveryCanReuseVerifiedRestoreStage()
    {
        File.WriteAllText(transaction.Target, "old bytes");
        transaction.Apply(transaction.Plan());
        File.Copy(transaction.Backup, Path.Combine(transaction.DirectoryPath, "restore.xlam"));
        Assert.Equal("rolled-back", transaction.Recover().Status);
        Assert.Equal("old bytes", File.ReadAllText(transaction.Target));
    }

    [Fact]
    public void InvalidJournalDoesNotModifyInstalledFile()
    {
        File.WriteAllText(transaction.Target, "old bytes");
        transaction.Apply(transaction.Plan());
        File.WriteAllText(Path.Combine(transaction.DirectoryPath, "transaction.json"),
            "{\"SchemaVersion\":9,\"State\":\"prepared\",\"SourceHash\":\"invalid\",\"TargetHash\":null}");
        Assert.Throws<IOException>(() => transaction.Recover());
        Assert.Equal("new bytes", File.ReadAllText(transaction.Target));
    }

    [Theory]
    [InlineData("staged", "blocked")]
    [InlineData("backed-up", "blocked")]
    [InlineData("prepared", "rolled-back")]
    [InlineData("replaced", "rolled-back")]
    public void FailuresAtEachPhasePreserveOrRestoreOriginal(string phase, string expectedStatus)
    {
        File.WriteAllText(transaction.Target, "old bytes");
        var result = transaction.Apply(transaction.Plan(), current =>
        {
            if (current == phase) throw new IOException("Injected failure");
        });
        Assert.Equal(expectedStatus, result.Status);
        Assert.Equal("old bytes", File.ReadAllText(transaction.Target));
    }

    [Theory]
    [InlineData("prepared")]
    [InlineData("replaced")]
    public void InterruptedTransactionCanBeRecoveredByNewInstance(string phase)
    {
        File.WriteAllText(transaction.Target, "old bytes");
        Assert.Throws<SimulatedInterruption>(() => transaction.Apply(transaction.Plan(), current =>
        {
            if (current == phase) throw new SimulatedInterruption();
        }));
        Assert.Equal("rolled-back", new TestFileTransaction(transaction.DirectoryPath).Recover().Status);
        Assert.Equal("old bytes", File.ReadAllText(transaction.Target));
    }

    [Fact]
    public void ChangedStagingFileIsRejectedBeforeReplacement()
    {
        File.WriteAllText(transaction.Target, "old bytes");
        var result = transaction.Apply(transaction.Plan(), phase =>
        {
            if (phase == "staged") File.WriteAllText(Path.Combine(transaction.DirectoryPath, "staged.xlam"), "tampered");
        });
        Assert.Equal("blocked", result.Status);
        Assert.Equal("old bytes", File.ReadAllText(transaction.Target));
    }

    [Fact]
    public void UnexpectedTargetAfterReplacementRequiresManualRecovery()
    {
        File.WriteAllText(transaction.Target, "old bytes");
        var result = transaction.Apply(transaction.Plan(), phase =>
        {
            if (phase == "replaced") File.WriteAllText(transaction.Target, "external change");
        });
        Assert.Equal("recovery-required", result.Status);
        Assert.Equal("external change", File.ReadAllText(transaction.Target));
        Assert.Equal("old bytes", File.ReadAllText(transaction.Backup));
    }

    [Fact]
    public void CorruptBackupPreventsDestructiveRecovery()
    {
        File.WriteAllText(transaction.Target, "old bytes");
        Assert.Equal("installed", transaction.Apply(transaction.Plan()).Status);
        File.WriteAllText(transaction.Backup, "corrupt");
        Assert.Throws<IOException>(() => transaction.Recover());
        Assert.Equal("new bytes", File.ReadAllText(transaction.Target));
    }

    [Fact]
    public void UnknownDirectoryCannotBeUsedForRecovery()
    {
        Assert.Throws<IOException>(() => new TestFileTransaction(Path.GetTempPath()).Recover());
    }

    [Fact]
    public void LaterChangesAfterRollbackArePreserved()
    {
        File.WriteAllText(transaction.Target, "old bytes");
        transaction.Apply(transaction.Plan());
        transaction.Recover();
        File.WriteAllText(transaction.Target, "new bytes");
        Assert.Throws<IOException>(() => transaction.Recover());
    }

    [Fact]
    public void RehearsalRejectsUnreleasedSourceBeforeCreatingOutput()
    {
        var output = Path.Combine(transaction.DirectoryPath, "unused");
        Assert.Throws<IOException>(() => InstallationRehearsal.Run(transaction.Source, output));
        Assert.False(Directory.Exists(output));
    }

    public void Dispose() => Directory.Delete(transaction.DirectoryPath, recursive: true);
}
