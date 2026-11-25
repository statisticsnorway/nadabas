namespace Nadabas.Setup;

// Test-only inputs. IT policy is not yet read from Windows or a management service.
internal sealed record UpdatePolicy(bool UserEnabled = true, bool? ItEnabled = null, string? PinnedVersion = null);

internal static class TestUpdateCoordinator
{
    internal static TransactionResult Apply(TestFileTransaction transaction, string version,
        UpdatePolicy policy, Func<bool> excelRunning, Action<string>? checkpoint = null)
    {
        if (!(policy.ItEnabled ?? policy.UserEnabled)) return new("disabled-by-policy");
        if (policy.PinnedVersion is not null && policy.PinnedVersion != version) return new("pinned-version-mismatch");
        if (excelRunning()) return new("deferred-excel-running");
        var plan = transaction.Plan();
        // Recheck before preparation and immediately before rename. This still
        // does not prevent Excel starting after the last process observation.
        if (excelRunning()) return new("deferred-excel-running");
        return transaction.Apply(plan, checkpoint, beforeReplace: () =>
        {
            try
            {
                if (excelRunning()) throw new IOException("Excel started during update preparation; replacement cancelled.");
            }
            catch (Exception ex) when (ex is not IOException)
            {
                throw new IOException("Could not confirm Excel is closed before replacement.", ex);
            }
        });
    }
}

internal sealed record UpdateRehearsalReport(int SchemaVersion, string Kind, string IncomingVersion,
    string PreviousPackage, string ExcelDetection, string PolicySource, bool SourceUnchanged,
    bool Passed, IReadOnlyList<RehearsalCase> Cases);

internal static class UpdateRehearsal
{
    internal static UpdateRehearsalReport Run(string source, string outputDirectory)
    {
        TestFileTransaction.CheckPath(source);
        TestFileTransaction.CheckPath(outputDirectory);
        var expected = ReleaseBaseline.RequireAddin(source);
        var cases = new List<RehearsalCase>();
        foreach (var scenario in new[] { "automatic-upgrade", "excel-open-then-closed", "it-disables-update",
            "it-pin-blocks-update", "it-pin-allows-update", "it-overrides-user-disabled", "user-disabled",
            "failure-after-replacement", "already-current", "excel-starts-before-apply" })
        {
            var transaction = TestFileTransaction.Create(outputDirectory);
            File.Copy(source, transaction.Source);
            if (TestFileTransaction.Hash(transaction.Source) != expected) throw new IOException("Source changed during preparation.");
            if (scenario == "already-current") File.Copy(transaction.Source, transaction.Target);
            else File.WriteAllText(transaction.Target, "NADABAS synthetic previous test package v1; not an Excel add-in or released version.");
            var before = TestFileTransaction.Hash(transaction.Target);
            var policy = scenario switch
            {
                "it-disables-update" => new UpdatePolicy(ItEnabled: false),
                "it-pin-blocks-update" => new UpdatePolicy(PinnedVersion: "synthetic-previous"),
                "it-pin-allows-update" => new UpdatePolicy(PinnedVersion: "6.02.005"),
                "it-overrides-user-disabled" => new UpdatePolicy(UserEnabled: false, ItEnabled: true),
                "user-disabled" => new UpdatePolicy(UserEnabled: false),
                _ => new UpdatePolicy()
            };
            var probes = 0;
            var result = TestUpdateCoordinator.Apply(transaction, "6.02.005", policy,
                () => scenario == "excel-open-then-closed" || (scenario == "excel-starts-before-apply" && ++probes == 2),
                phase => { if (scenario == "failure-after-replacement" && phase == "replaced") throw new IOException("Injected upgrade failure."); });
            var intermediatePassed = true;
            if (scenario == "excel-open-then-closed")
            {
                intermediatePassed = result.Status == "deferred-excel-running" &&
                    TestFileTransaction.Hash(transaction.Target) == before && !File.Exists(transaction.Backup) &&
                    !File.Exists(Path.Combine(transaction.DirectoryPath, "transaction.json"));
                cases.Add(new("excel-open-defers-without-writing", transaction.DirectoryPath, result.Status,
                    before, TestFileTransaction.Hash(transaction.Target), intermediatePassed));
                result = TestUpdateCoordinator.Apply(transaction, "6.02.005", policy, () => false);
            }
            var desiredStatus = scenario switch
            {
                "it-disables-update" or "user-disabled" => "disabled-by-policy",
                "it-pin-blocks-update" => "pinned-version-mismatch",
                "failure-after-replacement" => "rolled-back",
                "already-current" => "unchanged",
                "excel-starts-before-apply" => "deferred-excel-running",
                _ => "installed"
            };
            var after = TestFileTransaction.Hash(transaction.Target);
            var desiredHash = desiredStatus is "installed" or "unchanged" ? expected : before;
            var noWriteRequired = desiredStatus is "disabled-by-policy" or "pinned-version-mismatch" or "deferred-excel-running" or "unchanged";
            var backupPassed = noWriteRequired ? !File.Exists(transaction.Backup) &&
                !File.Exists(Path.Combine(transaction.DirectoryPath, "transaction.json")) : TestFileTransaction.Hash(transaction.Backup) == before;
            cases.Add(new(scenario, transaction.DirectoryPath, result.Status, before, after,
                intermediatePassed && result.Status == desiredStatus && after == desiredHash && backupPassed));
        }
        var unchanged = TestFileTransaction.Hash(source) == expected;
        return new(1, "update-rehearsal", "6.02.005", "synthetic-test-bytes-not-a-release",
            "simulated-no-excel-started", "in-memory-test-inputs", unchanged, unchanged && cases.All(c => c.Passed), cases);
    }
}
