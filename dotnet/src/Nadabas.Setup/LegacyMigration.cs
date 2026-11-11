using Microsoft.Win32;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace Nadabas.Setup;

public sealed record LegacyRegistration(string View, string Key, string Name, string Kind, string? Value,
    string? CandidatePath, string Evidence);
public sealed record LegacyCandidate(string Path, string Classification, string? Sha256,
    IReadOnlyList<string> FoundThrough, string ProposedAction);
public sealed record LegacyMigrationReport(int SchemaVersion, string Kind, bool CanApply,
    IReadOnlyList<LegacyRegistration> Registrations, IReadOnlyList<LegacyCandidate> Candidates,
    IReadOnlyList<string> ReviewItems, IReadOnlyList<string> RequiredChecks);

// Inventory only. No registry writes, Excel automation, file moves or deletion.
public static class LegacyMigration
{
    private const string BaselineHash = "2F7432B9B04C414337EB1AFEB07E2B93F05E4A5A6D151073FA967A2F7EA89A85";

    internal static bool IsCandidateName(string name) =>
        name.StartsWith("NADABAS.", StringComparison.OrdinalIgnoreCase) &&
        new[] { ".xlam", ".xla", ".xlsm" }.Contains(Path.GetExtension(name), StringComparer.OrdinalIgnoreCase);

    internal static string? RegistrationPath(string section, string name, string? value)
    {
        string text;
        if (section == "Add-in Manager") text = name;
        else
        {
            if (!Regex.IsMatch(name, "^OPEN[0-9]*$", RegexOptions.IgnoreCase) || value is null) return null;
            text = value.Trim();
            if (text.StartsWith("/R ", StringComparison.OrdinalIgnoreCase))
            {
                text = text[3..].TrimStart();
                if (!text.StartsWith('"')) return null; // Ambiguous command-like values need review.
            }
        }
        text = text.Trim();
        if (text.StartsWith('"') && text.EndsWith('"') && text.Length > 1) text = text[1..^1];
        if (text.Contains('"') || !Path.IsPathFullyQualified(text) || !IsCandidateName(Path.GetFileName(text))) return null;
        return Path.GetFullPath(text);
    }

    public static LegacyMigrationReport Collect(string addinDirectory)
    {
        if (!Path.IsPathFullyQualified(addinDirectory)) throw new ArgumentException("Use an absolute add-in directory.");
        var registrations = new List<LegacyRegistration>();
        var review = new List<string>();
        var files = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var directories = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            [Path.GetFullPath(addinDirectory)] = "user-addins-directory",
            [Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Excel", "XLSTART")] = "user-xlstart"
        };
        if (OperatingSystem.IsWindows()) ReadWindows(registrations, directories, review);
        else review.Add("Windows registry inventory is unavailable on this operating system.");
        void Found(string path, string evidence)
        {
            if (!files.TryGetValue(path, out var sources)) files[path] = sources = new(StringComparer.OrdinalIgnoreCase);
            sources.Add(evidence);
        }
        foreach (var registration in registrations)
            if (registration.CandidatePath is not null) Found(registration.CandidatePath, registration.Evidence);
        foreach (var (directory, source) in directories)
        {
            try
            {
                if (NeedsLocationReview(directory)) { review.Add("Startup/add-in location requires review; not traversed: " + directory); continue; }
                foreach (var file in Directory.EnumerateFiles(directory, "*", SearchOption.TopDirectoryOnly))
                    if (IsCandidateName(Path.GetFileName(file))) Found(file, source);
            }
            catch (DirectoryNotFoundException) { } // Optional startup folders may not exist.
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
            { review.Add("Directory inventory incomplete: " + directory + " (" + ex.GetType().Name + ")"); }
        }
        var managed = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NADABAS-Loader-Test");
        var candidates = files.OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
            .Select(pair => Inspect(pair.Key, pair.Value.ToArray(), managed)).ToArray();
        if (candidates.Count(c => c.Classification != "managed-test-payload") > 1)
            review.Add("Multiple legacy candidates found; determine which registration/workbook is actually used before switching.");
        return new(1, "legacy-migration-plan", false, registrations, candidates, review,
        [
            "Inventory is read-only and limited to this user's registrations and discovered startup locations; it is not proof of current Excel activation.",
            "Close Excel before migration; recheck exact registry values, paths, hashes, signatures and permissions.",
            "Back up identified legacy files and exact registry value names/types/contents, then verify backup hashes before switching.",
            "Install the new loader in its managed directory; deactivate only the exact legacy registrations after successful preparation.",
            "An XLSTART/alternate-startup copy needs a verified reversible move out of the auto-load location; removing registration alone is insufficient.",
            "Preserve unknown/custom/shared files and unrelated add-ins. Do not delete NADABAS folders, databases, workbooks or user settings.",
            "Verify a single NADABAS instance, version, ribbon and database operations; test rollback of registration and files without overwriting intervening user changes.",
            "Machine-wide, policy-managed, renamed add-ins and other user profiles require separate IT review; this tool does not scan or change them."
        ]);
    }

    internal static LegacyCandidate Inspect(string path, string[] sources, string managedRoot)
    {
        var normalized = Path.GetFullPath(path);
        if (normalized.StartsWith(Path.TrimEndingDirectorySeparator(managedRoot) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            return new(path, "managed-test-payload", null, sources, "Preserve; belongs to the new test installation, not the legacy migration.");
        try
        {
            if (NeedsLocationReview(path)) return new(path, "shared-or-linked-review", null, sources, "Preserve; do not overwrite a shared, mapped-network or linked location.");
            using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var hash = Convert.ToHexString(SHA256.HashData(file));
            var recognized = hash == BaselineHash;
            return new(path, recognized ? "known-6.02.005-bytes" : "unknown-or-custom-content", hash, sources,
                recognized ? "Candidate for verified backup and exact-registration handoff; no overwrite in the original location."
                    : "Preserve and review contents/version/signature before choosing migration behavior.");
        }
        catch (FileNotFoundException) { return new(path, "missing-file", null, sources, "Review stale registration; do not remove solely on a filename match."); }
        catch (DirectoryNotFoundException) { return new(path, "missing-file", null, sources, "Review stale registration; location may be unavailable."); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
        { return new(path, "inspection-failed", null, sources, "Preserve; inspection failed: " + ex.GetType().Name); }
    }

    private static bool NeedsLocationReview(string path)
    {
        if (!Path.IsPathFullyQualified(path) || path.StartsWith(@"\\", StringComparison.Ordinal)) return true;
        if (OperatingSystem.IsWindows() && new DriveInfo(Path.GetPathRoot(path)!).DriveType == DriveType.Network) return true;
        var ancestors = new Stack<string>();
        for (var item = Path.GetFullPath(path); !string.IsNullOrEmpty(item); item = Path.GetDirectoryName(item)) ancestors.Push(item);
        foreach (var item in ancestors)
        {
            try { if ((File.GetAttributes(item) & FileAttributes.ReparsePoint) != 0) return true; }
            catch (FileNotFoundException) { break; }
            catch (DirectoryNotFoundException) { break; }
        }
        return false;
    }

    [SupportedOSPlatform("windows")]
    private static void ReadWindows(List<LegacyRegistration> registrations, Dictionary<string, string> directories, List<string> review)
    {
        foreach (var view in new[] { RegistryView.Registry32, RegistryView.Registry64 })
        {
            using var user = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
            foreach (var version in new[] { "12.0", "14.0", "15.0", "16.0" })
                foreach (var section in new[] { "Options", "Add-in Manager" })
                {
                    var path = $@"Software\Microsoft\Office\{version}\Excel\{section}";
                    try
                    {
                        using var key = user.OpenSubKey(path, writable: false);
                        if (key is null) continue;
                        if (section == "Options" && key.GetValue("AltStartup") is string alternate && !string.IsNullOrWhiteSpace(alternate))
                            directories[alternate] = "alternate-startup";
                        foreach (var name in key.GetValueNames())
                        {
                            if (section == "Options" && !name.StartsWith("OPEN", StringComparison.OrdinalIgnoreCase)) continue;
                            var value = key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames) as string;
                            if (!name.Contains("NADABAS", StringComparison.OrdinalIgnoreCase) && !(value?.Contains("NADABAS", StringComparison.OrdinalIgnoreCase) ?? false)) continue;
                            var candidate = RegistrationPath(section, name, value);
                            var evidence = section == "Options" ? "excel-open-registration-not-runtime-proof" : "addin-manager-entry-not-activation-proof";
                            registrations.Add(new(view.ToString(), path, name, key.GetValueKind(name).ToString(), value, candidate, evidence));
                            if (candidate is null) review.Add("NADABAS registration needs manual interpretation: " + view + " " + path + " / " + name);
                        }
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException or ArgumentException)
                    { review.Add("Registry inventory incomplete: " + path + " (" + ex.GetType().Name + ")"); }
                }
            foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
            {
                try
                {
                    using var root = RegistryKey.OpenBaseKey(hive, view);
                    using var key = root.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe", writable: false);
                    if (key?.GetValue(null) is string executable && Path.IsPathFullyQualified(executable))
                        directories[Path.Combine(Path.GetDirectoryName(executable)!, "XLSTART")] = "office-xlstart";
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
                { review.Add("Office startup directory could not be inventoried: " + hive + " " + view); }
            }
        }
    }
}
