using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;
using Microsoft.Win32;

namespace Nadabas.Setup;

public sealed record Finding(string Id, string Status, string Evidence, string? Value = null);
public sealed record DiagnosticReport(int SchemaVersion, string ToolVersion, string OperatingSystem,
    string ProcessArchitecture, bool Complete, IReadOnlyList<Finding> Findings);

public static class Diagnostics
{
    public static DiagnosticReport Collect(string addinDirectory)
    {
        var findings = new List<Finding>(InspectAddins(addinDirectory));
        if (OperatingSystem.IsWindows())
            InspectWindows(findings);
        else
            findings.Add(new("windows.registry", "unsupported", "Registry inventory requires Windows."));

        findings.Add(new("excel.runtime", "not-tested",
            "Excel was not started. Confirm Office version, bitness, add-in activation and macro trust in Excel."));
        findings.Add(new("database.connection", "not-tested",
            "No database or credential settings were read. Driver registration does not prove connectivity."));
        return new(1, "0.1.0", RuntimeInformation.OSDescription,
            RuntimeInformation.ProcessArchitecture.ToString(),
            !findings.Any(f => f.Status is "error" or "unsupported"), findings);
    }

    public static IReadOnlyList<Finding> InspectAddins(string directory)
    {
        var findings = new List<Finding>();
        try
        {
            // Enumerate directly: Directory.Exists would hide access-denied errors.
            var files = Directory.EnumerateFiles(directory, "*", SearchOption.TopDirectoryOnly)
                .Where(path => IsNadabasAddin(Path.GetFileName(path)))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToArray();
            findings.Add(new("addin.directory", "observed", "Only this directory was inspected; custom Excel locations may differ.", directory));
            if (files.Length == 0)
                findings.Add(new("addin.files", "not-found", "No NADABAS XLAM candidates in the inspected directory."));
            foreach (var file in files)
                findings.Add(new("addin.file", "observed", "Filename match only; contents, version, signature and activation were not verified.", file));
            if (files.Length > 1)
                findings.Add(new("addin.multiple", "warning", "Multiple candidates found. Review in Excel before any future installation."));
        }
        catch (DirectoryNotFoundException)
        {
            findings.Add(new("addin.directory", "not-found", "The inspected directory does not exist.", directory));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or SecurityException)
        {
            findings.Add(new("addin.directory", "error", "Directory inspection failed: " + ex.GetType().Name, directory));
        }
        return findings;
    }

    public static bool IsNadabasAddin(string name) =>
        name.Equals("NADABAS.xlam", StringComparison.OrdinalIgnoreCase) ||
        (name.StartsWith("NADABAS.", StringComparison.OrdinalIgnoreCase) &&
         name.EndsWith(".xlam", StringComparison.OrdinalIgnoreCase));

    [SupportedOSPlatform("windows")]
    private static void InspectWindows(List<Finding> findings)
    {
        RegistryView[] views = Environment.Is64BitOperatingSystem
            ? [RegistryView.Registry32, RegistryView.Registry64]
            : [RegistryView.Registry32];
        foreach (var view in views)
        {
            foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
            {
                ReadRegistry(findings, $"excel.app-path.{hive}.{view}", hive, view,
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
                    "Excel App Paths registration; not proof of a working installation or Office bitness.");
            }
            foreach (var version in new[] { "12.0", "16.0" })
            {
                ReadRegistry(findings, $"ace.{version}.{view}", RegistryHive.ClassesRoot, view,
                    $@"Microsoft.ACE.OLEDB.{version}\CLSID",
                    "ACE ProgID registration in this registry view; provider loading was not tested.");
            }
        }
    }

    [SupportedOSPlatform("windows")]
    private static void ReadRegistry(List<Finding> findings, string id, RegistryHive hive,
        RegistryView view, string path, string evidence)
    {
        try
        {
            using var root = RegistryKey.OpenBaseKey(hive, view);
            using var key = root.OpenSubKey(path, writable: false);
            var value = key?.GetValue(null) as string;
            findings.Add(new(id, string.IsNullOrWhiteSpace(value) ? "not-found" : "observed", evidence, value));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or SecurityException)
        {
            findings.Add(new(id, "error", "Registry inspection failed: " + ex.GetType().Name));
        }
    }
}
