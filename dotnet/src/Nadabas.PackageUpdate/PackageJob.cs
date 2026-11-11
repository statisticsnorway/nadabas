using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text.Json;

namespace Nadabas.PackageUpdate;

internal sealed record ApplyResult(string Status, string? Version = null);

internal sealed record PendingPackage(int SchemaVersion, string Root, string PreviousVersion,
    string TargetVersion, string Sha256, bool Smoke);

[SupportedOSPlatform("windows")]
internal static class PackageJob
{
    private const string Marker = "NADABAS version-store prototype v1";
    private const string JobName = "package-update.json";

    internal static PendingPackage Stage(string root, string installer, string cache, bool smoke)
    {
        root = PackageRules.SafePath(root);
        installer = PackageRules.SafePath(installer);
        cache = PackageRules.SafePath(cache);
        PackageRules.RequireOutside(root, cache);
        if (Directory.Exists(cache) || File.Exists(cache)) throw new IOException("Use a new external cache directory.");
        var current = InstalledVersion(root, smoke);
        var incoming = VerifyPackage(installer, smoke);
        PackageRules.RequireUpgrade(current, incoming);
        if (!smoke) UpdatePolicy.Read().RequireAllowed(true, incoming);
        var self = Environment.ProcessPath ?? throw new IOException("Cannot identify updater executable.");
        if (!Path.GetFileName(self).Equals("Nadabas.PackageUpdate.exe", StringComparison.OrdinalIgnoreCase))
            throw new IOException("Run the signed published single-file updater.");
        PublisherTrust.Verify(self);
        Directory.CreateDirectory(cache);
        var cachedInstaller = Path.Combine(cache, "setup.exe");
        File.Copy(installer, cachedInstaller);
        var verified = VerifyPackage(cachedInstaller, smoke);
        if (verified != incoming) throw new IOException("Package changed during preparation.");
        File.Copy(self, Path.Combine(cache, "Nadabas.PackageUpdate.exe"));
        var pending = new PendingPackage(1, root, current, incoming, Hash(cachedInstaller), smoke);
        Write(cache, JobName, pending);
        Write(cache, "status.json", new { status = "ready", pending.TargetVersion });
        return pending;
    }

    internal static async Task<ApplyResult> Apply(string cache)
    {
        cache = PackageRules.SafePath(cache);
        var jobPath = PackageRules.SafePath(Path.Combine(cache, JobName));
        if (new FileInfo(jobPath).Length > 8192) throw new IOException("Invalid job size.");
        var pending = JsonSerializer.Deserialize<PendingPackage>(File.ReadAllText(jobPath)) ?? throw new IOException("Missing job.");
        if (pending.SchemaVersion != 1) throw new IOException("Unsupported job schema.");
        var root = PackageRules.SafePath(pending.Root);
        PackageRules.RequireOutside(root, cache);
        PackageRules.RequireOutside(root, Environment.ProcessPath!);
        using var gate = new FileStream(Path.Combine(cache, "apply.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
        if (File.Exists(PackageRules.SafePath(Path.Combine(cache, "cancelled.json")))) throw new IOException("This update was cancelled.");
        if (File.Exists(Path.Combine(cache, "apply-started.json")))
            throw new IOException("This job has already started. Inspect the result and repair before retrying; automatic replay is disabled.");
        if (AnyExcel()) return new ApplyResult("deferred-excel-running");
        var current = InstalledVersion(root, pending.Smoke);
        if (PackageRules.Normalize(current) != PackageRules.Normalize(pending.PreviousVersion))
            throw new IOException("Installed package changed after staging.");
        PackageRules.RequireUpgrade(current, pending.TargetVersion);
        if (!pending.Smoke) UpdatePolicy.Read().RequireAllowed(true, pending.TargetVersion);
        var installer = PackageRules.SafePath(Path.Combine(cache, "setup.exe"));
        // Retain a Windows read lease to prevent replacement between validation and execution.
        using var lease = new FileStream(installer, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (Hash(installer) != pending.Sha256 || VerifyPackage(installer, pending.Smoke) != pending.TargetVersion)
            throw new IOException("Cached package differs from the verified candidate.");
        if (AnyExcel()) return new ApplyResult("deferred-excel-running");
        if (!pending.Smoke) UpdatePolicy.Read().RequireAllowed(true, pending.TargetVersion);
        Write(cache, "apply-started.json", new { pending.TargetVersion, startedUtc = DateTimeOffset.UtcNow });
        try
        {
            var installed = await RunSetup(root, pending.Smoke, installer, pending.TargetVersion, Path.Combine(cache, "installer.log"));
            var result = new { status = "installed", version = installed };
            Write(cache, "status.json", result);
            return new ApplyResult("installed", installed);
        }
        catch (Exception ex)
        {
            Write(cache, "status.json", new { status = "repair-required", error = ex.Message });
            throw;
        }
    }

    internal static async Task<string> RunSetup(string root, bool smoke, string installer, string targetVersion, string log)
    {
        var command = new ProcessStartInfo(installer) { UseShellExecute = false, CreateNoWindow = true };
        foreach (var argument in new[] { "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/DIR=" + root, "/LOG=" + log })
            command.ArgumentList.Add(argument);
        using var process = Process.Start(command) ?? throw new IOException("Installer could not start.");
        await process.WaitForExitAsync();
        if (process.ExitCode != 0) throw new IOException("Setup failed with exit code " + process.ExitCode + ". Keep the installation log.");
        var installed = InstalledVersion(root, smoke);
        if (PackageRules.Normalize(installed) != PackageRules.Normalize(targetVersion))
            throw new IOException("Post-install version did not match. Repair is required.");
        return installed;
    }

    internal static string InstalledVersion(string root, bool smoke)
    {
        if (File.ReadAllText(PackageRules.SafePath(Path.Combine(root, ".nadabas-version-store"))).TrimEnd() != Marker)
            throw new IOException("Unknown installation.");
        using var user = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64);
        var identity = smoke ? "NadabasLoaderInstallerSmoke_is1" : "{8C18DA56-6F3D-4B42-B114-CC449364E8F9}_is1";
        using var key = user.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + identity);
        var location = key?.GetValue("InstallLocation") as string;
        if (location is null || !PackageRules.SafePath(location).Equals(root, StringComparison.OrdinalIgnoreCase))
            throw new IOException("Registered installation does not match the requested target.");
        return key?.GetValue("DisplayVersion") as string ?? throw new IOException("Installed version missing.");
    }

    internal static string VerifyPackage(string path, bool smoke)
    {
        PublisherTrust.Verify(path);
        var info = FileVersionInfo.GetVersionInfo(path);
        if (info.ProductName?.Trim() != (smoke ? "NADABAS Installer Smoke" : "NADABAS Loader Test"))
            throw new IOException("Signed executable is not the expected NADABAS installer.");
        return new Version(info.FileMajorPart, info.FileMinorPart, info.FileBuildPart, info.FilePrivatePart).ToString();
    }

    internal static bool AnyExcel()
    {
        var processes = Process.GetProcessesByName("EXCEL");
        try { return processes.Length != 0; }
        finally { foreach (var process in processes) process.Dispose(); }
    }

    private static string Hash(string path)
    {
        using var input = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(input));
    }

    internal static void Write<T>(string cache, string name, T value)
    {
        var target = PackageRules.SafePath(Path.Combine(cache, name));
        var temporary = Path.Combine(cache, Guid.NewGuid().ToString("N") + ".pending");
        File.WriteAllText(temporary, JsonSerializer.Serialize(value, new JsonSerializerOptions { WriteIndented = true }));
        File.Move(temporary, target, true);
    }
}
