using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text.Json;

namespace Nadabas.Setup;

internal sealed record ExcelStartupReport(int SchemaVersion, string Kind, string Status, string TestAddin,
    string? ExcelVersion, string? NadabasVersion, bool StartupCompleted, bool CopyUnchanged,
    bool SourceUnchanged, bool SessionClosed, bool AddinActivated, bool RegistrationCleaned,
    string RibbonVerification, IReadOnlyList<string> Errors);

[SupportedOSPlatform("windows")]
internal static class ExcelStartupTest
{
    internal static ExcelStartupReport Run(string source, string outputDirectory, bool interactive, TextWriter progress)
    {
        var expectedHash = ReleaseBaseline.RequireAddin(source);
        var processes = Process.GetProcessesByName("EXCEL");
        try
        {
            if (processes.Length != 0) throw new IOException("Close Excel after saving your work before running this test. No running Excel session was changed.");
        }
        finally { foreach (var process in processes) process.Dispose(); }
        TestFileTransaction.CheckPath(outputDirectory);
        Directory.CreateDirectory(outputDirectory);
        var transaction = TestFileTransaction.Create(outputDirectory);
        File.Copy(source, transaction.Source);
        if (TestFileTransaction.Hash(transaction.Source) != expectedHash) throw new IOException("Source changed while preparing test.");
        if (transaction.Apply(transaction.Plan()).Status != "installed") throw new IOException("Could not prepare test add-in.");
        var testAddin = Path.Combine(transaction.DirectoryPath, "NADABAS.TEST." + Guid.NewGuid().ToString("N") + ".xlam");
        File.Copy(transaction.Target, testAddin);
        File.SetAttributes(testAddin, FileAttributes.ReadOnly);

        dynamic? excel = null;
        dynamic? books = null;
        dynamic? addin = null;
        dynamic? blank = null;
        dynamic? addins = null;
        dynamic? registration = null;
        string? excelVersion = null;
        string? nadabasVersion = null;
        var startupCompleted = false;
        var sessionClosed = false;
        var addinActivated = false;
        var registrationCleaned = false;
        var errors = new List<string>();
        try
        {
            var type = Type.GetTypeFromProgID("Excel.Application") ?? throw new IOException("Desktop Excel is not registered.");
            excel = Activator.CreateInstance(type) ?? throw new IOException("Could not create Excel instance.");
            excel.AutomationSecurity = 2; // ByUI: respect existing Office macro policy, never force-enable macros.
            excel.EnableEvents = false;
            excel.AskToUpdateLinks = false;
            excel.Visible = true;
            excelVersion = Convert.ToString(excel.Version);
            books = excel.Workbooks;
            if ((int)books.Count != 0) throw new IOException("Unexpected startup workbooks; no test add-in was opened.");
            addins = excel.AddIns;
            for (var index = 1; index <= (int)addins.Count; index++)
            {
                dynamic existing = addins.Item(index);
                try
                {
                    if (Diagnostics.IsNadabasAddin(Path.GetFileName(Convert.ToString(existing.FullName)!)) && (bool)existing.Installed)
                        throw new IOException("Another NADABAS add-in is active. Deactivate it yourself before testing; its state was not changed.");
                }
                finally { Marshal.FinalReleaseComObject(existing); }
            }
            blank = books.Add();
            excel.Caption = "NADABAS TEST 6.02.005 - " + Path.GetFileName(testAddin);
            excel.EnableEvents = true;
            registration = addins.Add(testAddin, false);
            if (!string.Equals(Path.GetFullPath(Convert.ToString(registration.FullName)!), testAddin, StringComparison.OrdinalIgnoreCase))
                throw new IOException("Excel registered a different file. Activation was not attempted.");
            registration.Installed = true;
            addinActivated = (bool)registration.Installed;
            if (!addinActivated) throw new IOException("Excel did not activate the test add-in.");
            addin = books.Item(Path.GetFileName(testAddin));
            if (!string.Equals(Path.GetFullPath(Convert.ToString(addin.FullName)!), testAddin, StringComparison.OrdinalIgnoreCase))
                throw new IOException("Excel did not open the expected test copy.");
            if (!(bool)addin.IsAddin) throw new IOException("The test file is not an Excel add-in.");
            excel.EnableEvents = true;
            addin.RunAutoMacros(1);
            startupCompleted = true;
            var macroPrefix = "'" + Path.GetFileName(testAddin) + "'!";
            excel.Run(macroPrefix + "InterfaceVersionUpdate.CancelScheduledVersionCheck");
            nadabasVersion = Convert.ToString(excel.Run(macroPrefix + "InterfaceVersionUpdate.InstalledVersion"));
            if (nadabasVersion != "6.02.005") throw new IOException("Runtime NADABAS version differs from the selected release: " + nadabasVersion);
            if (interactive)
            {
                progress.WriteLine("NADABAS " + nadabasVersion + " started from the test copy. Inspect Ribbon/About; do not open databases in this startup-only session.");
                progress.WriteLine("Close dialogs and press Enter here to end the test without saving.");
                progress.WriteLine("ACTIVE TEST FILE: " + testAddin);
                Console.ReadLine();
            }
        }
        catch (Exception ex) { errors.Add(ex.GetType().Name + ": " + ex.Message); }
        finally
        {
            if (addin is not null)
            {
                if (startupCompleted)
                {
                    try { addin.RunAutoMacros(2); }
                    catch (Exception ex) { errors.Add("AutoClose: " + ex.Message); }
                }
            }
            if (registration is not null)
            {
                try
                {
                    if (string.Equals(Convert.ToString(registration.FullName), testAddin, StringComparison.OrdinalIgnoreCase))
                        registration.Installed = false;
                }
                catch (Exception ex) { errors.Add("Deactivate test add-in: " + ex.Message); }
            }
            if (blank is not null)
            {
                try { blank.Close(false); }
                catch (Exception ex) { errors.Add("Close blank workbook: " + ex.Message); }
            }
            if (excel is not null)
            {
                try
                {
                    if (books is not null && (int)books.Count != 0)
                        errors.Add("Other workbooks are open. Excel was left running for you to close manually.");
                    else { excel.Quit(); sessionClosed = true; }
                }
                catch (Exception ex) { errors.Add("Quit Excel: " + ex.Message); }
            }
            foreach (object? com in new object?[] { addin, registration, addins, blank, books, excel })
                if (com is not null && Marshal.IsComObject(com)) Marshal.FinalReleaseComObject(com);
        }
        if (sessionClosed && excelVersion is not null)
        {
            try { TestAddinRegistration.RemoveTestEntry(excelVersion, testAddin); registrationCleaned = true; }
            catch (Exception ex) { errors.Add("Remove test registration: " + ex.Message); }
        }
        var copyUnchanged = TestFileTransaction.Hash(testAddin) == expectedHash;
        var sourceUnchanged = TestFileTransaction.Hash(source) == expectedHash;
        var report = new ExcelStartupReport(2, "excel-startup-test",
            errors.Count == 0 && copyUnchanged && sourceUnchanged && startupCompleted && sessionClosed && addinActivated && registrationCleaned ? "activation-verified-ribbon-pending" : "needs-review",
            testAddin, excelVersion, nadabasVersion, startupCompleted, copyUnchanged, sourceUnchanged, sessionClosed,
            addinActivated, registrationCleaned, "requires-user-observation", errors);
        File.WriteAllText(Path.Combine(transaction.DirectoryPath, "excel-startup-report.json"),
            JsonSerializer.Serialize(report, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, WriteIndented = true }));
        return report;
    }
}
