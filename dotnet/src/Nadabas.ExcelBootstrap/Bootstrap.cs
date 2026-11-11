using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using ExcelDna.Integration;
using Nadabas.Versioning;

namespace Nadabas.ExcelBootstrap;

// Test-only loader; sign the packed XLL before distribution. Register only this XLL in a
// profile/test. No network, registry edits, or changes to Office trust settings.
public sealed class Bootstrap : IExcelAddIn
{
    private string? root;
    private VersionSelection? selection;
    private dynamic? workbook;
    private bool started;
    private string? runtimeVersion;
    private string updateStatus = "none";
    private string? updateError;
    private UpdateNotifications? notifications;
    private bool managedVbaChecks;

    public void AutoOpen()
    {
        try
        {
            root = Path.GetDirectoryName((string)XlCall.Excel(XlCall.xlGetName))!;
            var startup = new VersionStore(root).SelectForStartup();
            selection = startup.Selection;
            updateStatus = startup.UpdateStatus;
            updateError = startup.UpdateError;
            WriteState("selected");
            ExcelAsyncUtil.QueueAsMacro(LoadSelectedVersion);
        }
        catch (Exception ex) { WriteState("failed", ex.Message); }
    }

    private void LoadSelectedVersion()
    {
        try
        {
            if (selection is null || started) return;
            dynamic excel = ExcelDnaUtil.Application;
            dynamic addins = excel.AddIns;
            try
            {
                for (var i = 1; i <= (int)addins.Count; i++)
                {
                    dynamic existing = addins.Item(i);
                    try
                    {
                        var name = Path.GetFileName((string)existing.FullName);
                        if (name.StartsWith("NADABAS", StringComparison.OrdinalIgnoreCase) &&
                            name.EndsWith(".xlam", StringComparison.OrdinalIgnoreCase) && (bool)existing.Installed)
                            throw new IOException("Another NADABAS XLAM is active; loader did not replace it.");
                    }
                    finally { Marshal.ReleaseComObject(existing); }
                }
            }
            finally { Marshal.ReleaseComObject(addins); }
            dynamic books = excel.Workbooks;
            // AddIns.Installed persists the payload independently of this loader.
            // Open read-only for this session instead. AutoClose is not called on
            // normal Excel shutdown, so correctness must not depend on deregistration.
            try { workbook = books.Open(selection.AddinPath, UpdateLinks: 0, ReadOnly: true, AddToMru: false); }
            finally { Marshal.ReleaseComObject(books); }
            if (!string.Equals((string)workbook.FullName, selection.AddinPath, StringComparison.OrdinalIgnoreCase) || !(bool)workbook.IsAddin)
                throw new IOException("Loaded workbook differs from selected version.");
            workbook.RunAutoMacros(1);
            started = true;
            var prefix = "'" + Path.GetFileName(selection.AddinPath) + "'!";
            excel.Run(prefix + "InterfaceVersionUpdate.CancelScheduledVersionCheck");
            try
            {
                excel.Run(prefix + "InterfaceVersionUpdate.EnableManagedUpdates");
                managedVbaChecks = true;
            }
            catch (COMException) { managedVbaChecks = false; } // Historical signed payloads have no bridge.
            runtimeVersion = Convert.ToString(excel.Run(prefix + "InterfaceVersionUpdate.InstalledVersion"));
            if (runtimeVersion != selection.ReleaseVersion) throw new IOException("Runtime NADABAS version does not match the selected approved release.");
            notifications = new UpdateNotifications(root!);
            notifications.Start();
            WriteState("loaded");
        }
        catch (Exception ex) { WriteState("failed", ex.Message); }
    }

    public void AutoClose()
    {
        notifications?.Dispose();
        notifications = null;
        string? error = null;
        try
        {
            if (started && workbook is not null) workbook.RunAutoMacros(2);
        }
        catch (Exception ex) { error = ex.Message; }
        try
        {
            if (workbook is not null && selection is not null &&
                string.Equals((string)workbook.FullName, selection.AddinPath, StringComparison.OrdinalIgnoreCase))
                workbook.Close(SaveChanges: false);
        }
        catch (Exception ex) { error = (error is null ? "" : error + "; ") + ex.Message; }
        finally
        {
            if (workbook is not null) Marshal.ReleaseComObject(workbook);
            selection?.Dispose();
            workbook = null;
            started = false;
        }
        WriteState(error is null ? "closed" : "cleanup-failed", error);
    }

    private void WriteState(string status, string? error = null)
    {
        if (root is null) return;
        var sessions = Path.Combine(root, "sessions");
        Directory.CreateDirectory(sessions);
        var path = Path.Combine(sessions, Process.GetCurrentProcess().Id + ".json");
        var pending = path + ".pending";
        File.WriteAllText(pending, JsonSerializer.Serialize(new
        {
            schemaVersion = 1,
            loaderVersion = typeof(Bootstrap).Assembly.GetName().Version?.ToString(3) ?? "unknown",
            loadingMode = "session-workbook-no-payload-registration",
            kind = "version-loader-session",
            status,
            processId = Process.GetCurrentProcess().Id,
            slot = selection?.Slot,
            expectedReleaseVersion = selection?.ReleaseVersion,
            addinPath = selection?.AddinPath,
            runtimeVersion,
            updateStatus,
            updateError,
            managedVbaChecks,
            ribbonVerification = "requires-user-observation",
            error
        }, new JsonSerializerOptions { WriteIndented = true }));
        File.Move(pending, path, true);
    }
}
