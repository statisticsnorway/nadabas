using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Nadabas.Setup;

internal static class ExcelProcessProbe
{
    internal static bool IsRunning()
    {
        if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("Excel detection requires Windows.");
        return HasProcesses(() => Process.GetProcessesByName("EXCEL"));
    }

    internal static bool HasProcesses(Func<Process[]> enumerate)
    {
        // Enumeration failures propagate: uncertainty must never mean Excel is closed.
        var processes = enumerate();
        try { return processes.Length != 0; }
        finally { foreach (var process in processes) process.Dispose(); }
    }
}

internal sealed record LiveUpdateReport(int SchemaVersion, string Kind, bool Passed,
    string Workspace, string ExcelDetection, string PolicySource, bool SessionClosed,
    bool SourceUnchanged, IReadOnlyList<RehearsalCase> Cases, IReadOnlyList<string> Errors);

[SupportedOSPlatform("windows")]
internal static class LiveUpdateRehearsal
{
    internal static LiveUpdateReport Run(string source, string outputDirectory)
    {
        TestFileTransaction.CheckPath(source);
        TestFileTransaction.CheckPath(outputDirectory);
        var expected = ReleaseBaseline.RequireAddin(source);
        if (ExcelProcessProbe.IsRunning()) throw new IOException("Save and close Excel first. Existing sessions were not changed.");
        var transaction = TestFileTransaction.Create(outputDirectory);
        File.Copy(source, transaction.Source);
        if (TestFileTransaction.Hash(transaction.Source) != expected) throw new IOException("Source changed during preparation.");
        File.WriteAllText(transaction.Target, "Synthetic previous test package; not a released Excel add-in.");
        var before = TestFileTransaction.Hash(transaction.Target);
        var cases = new List<RehearsalCase>();
        var errors = new List<string>();
        dynamic? excel = null;
        dynamic? books = null;
        var closed = false;
        try
        {
            var type = Type.GetTypeFromProgID("Excel.Application") ?? throw new IOException("Desktop Excel is not registered.");
            excel = Activator.CreateInstance(type) ?? throw new IOException("Could not create Excel test instance.");
            excel.AutomationSecurity = 3; // No workbook macros needed for process detection.
            excel.EnableEvents = false;
            books = excel.Workbooks;
            if ((int)books.Count != 0) throw new IOException("Unexpected startup workbooks; no update attempted.");
            var detected = ExcelProcessProbe.IsRunning();
            if (!detected) throw new IOException("Created Excel instance was not detected; no update attempted.");
            var result = TestUpdateCoordinator.Apply(transaction, "6.02.005", new(), ExcelProcessProbe.IsRunning);
            var unchanged = TestFileTransaction.Hash(transaction.Target) == before;
            cases.Add(new("real-excel-defers-update", transaction.DirectoryPath, result.Status, before,
                TestFileTransaction.Hash(transaction.Target), result.Status == "deferred-excel-running" && unchanged &&
                !File.Exists(transaction.Backup) && !File.Exists(Path.Combine(transaction.DirectoryPath, "transaction.json"))));
        }
        catch (Exception ex) { errors.Add(ex.GetType().Name + ": " + ex.Message); }
        finally
        {
            if (excel is not null)
            {
                try
                {
                    // Only quit the instance we created, and only if no workbook needs preserving.
                    if (books is not null && (int)books.Count == 0) { excel.Quit(); closed = true; }
                    else errors.Add("Excel contains unexpected workbooks; left open for user inspection.");
                }
                catch (Exception ex) { errors.Add("Excel cleanup: " + ex.Message); }
            }
            if (books is not null) Marshal.FinalReleaseComObject(books);
            if (excel is not null) Marshal.FinalReleaseComObject(excel);
        }

        if (closed && errors.Count == 0 && cases.Count == 1 && cases[0].Passed)
        {
            try
            {
                var timeout = Stopwatch.StartNew();
                while (ExcelProcessProbe.IsRunning() && timeout.Elapsed < TimeSpan.FromSeconds(10)) Thread.Sleep(100);
                var result = TestUpdateCoordinator.Apply(transaction, "6.02.005", new(), ExcelProcessProbe.IsRunning);
                cases.Add(new("real-excel-closed-retry", transaction.DirectoryPath, result.Status, before,
                    TestFileTransaction.Hash(transaction.Target), result.Status == "installed" &&
                    TestFileTransaction.Hash(transaction.Target) == expected && TestFileTransaction.Hash(transaction.Backup) == before));
                if (result.Status == "installed")
                {
                    var rollback = transaction.Recover();
                    cases.Add(new("restore-after-live-test", transaction.DirectoryPath, rollback.Status, expected,
                        TestFileTransaction.Hash(transaction.Target), rollback.Status == "rolled-back" && TestFileTransaction.Hash(transaction.Target) == before));
                }
            }
            catch (Exception ex) { errors.Add(ex.GetType().Name + ": " + ex.Message); }
        }
        var sourceUnchanged = TestFileTransaction.Hash(source) == expected;
        return new(1, "live-update-rehearsal", errors.Count == 0 && closed && sourceUnchanged && cases.Count == 3 && cases.All(c => c.Passed),
            transaction.DirectoryPath, "real-windows-excel-processes", "in-memory-default-test-policy", closed, sourceUnchanged, cases, errors);
    }
}
