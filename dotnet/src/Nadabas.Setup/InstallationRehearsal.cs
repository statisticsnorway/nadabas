using System.Text.Json;

namespace Nadabas.Setup;

internal sealed record RehearsalCase(string Name, string Workspace, string Status,
    string? BeforeHash, string? AfterHash, bool Passed);
internal sealed record RehearsalReport(int SchemaVersion, string Kind, string ReleaseVersion,
    bool Passed, IReadOnlyList<RehearsalCase> Cases);
internal sealed class SimulatedInterruption : Exception;

internal static class InstallationRehearsal
{
    internal static RehearsalReport Run(string source, string outputDirectory)
    {
        TestFileTransaction.CheckPath(source);
        TestFileTransaction.CheckPath(outputDirectory);
        var expected = ReleaseBaseline.RequireAddin(source);
        Directory.CreateDirectory(outputDirectory);
        var cases = new List<RehearsalCase>();
        foreach (var scenario in new[] { "fresh-install", "identical", "replace-and-restore", "failure-after-replace", "interrupted-and-recovered" })
        {
            var transaction = TestFileTransaction.Create(outputDirectory);
            File.Copy(source, transaction.Source);
            if (TestFileTransaction.Hash(transaction.Source) != expected) throw new IOException("Source changed while preparing rehearsal.");
            if (scenario == "identical") File.Copy(transaction.Source, transaction.Target);
            else if (scenario != "fresh-install")
                File.WriteAllText(transaction.Target, "Synthetic previous file for NADABAS transaction testing. Not a released add-in.");

            var plan = transaction.Plan();
            TransactionResult result;
            try
            {
                result = transaction.Apply(plan, phase =>
                {
                    if (phase != "replaced") return;
                    if (scenario == "failure-after-replace") throw new IOException("Injected test failure.");
                    if (scenario == "interrupted-and-recovered") throw new SimulatedInterruption();
                });
            }
            catch (SimulatedInterruption)
            {
                // New object has no in-memory transaction state: recovery reads the durable journal.
                result = new TestFileTransaction(transaction.DirectoryPath).Recover();
            }
            var applyPassed = scenario == "identical" ? result.Status == "unchanged" :
                scenario is "failure-after-replace" or "interrupted-and-recovered" ? result.Status == "rolled-back" :
                result.Status == "installed" && TestFileTransaction.Hash(transaction.Target) == expected;
            if (scenario == "replace-and-restore" && applyPassed) result = transaction.Recover();
            var after = TestFileTransaction.Hash(transaction.Target);
            var desired = scenario is "fresh-install" or "identical" ? expected : plan.TargetHash;
            cases.Add(new(scenario, transaction.DirectoryPath, result.Status, plan.TargetHash, after, applyPassed && after == desired));
        }
        return new(1, "installation-rehearsal", "6.02.005", cases.All(c => c.Passed), cases);
    }
}
