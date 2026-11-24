using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text.Json;

namespace Nadabas.PackageUpdate;

internal sealed record RuntimeCleanupResult(string Status, string Architecture, int Files, long Bytes);

// Explicit maintenance only: retain packages, helpers, job records and recovery diagnostics.
[SupportedOSPlatform("windows")]
internal static class RuntimeCleanup
{
    internal static RuntimeCleanupResult Run(string root, string id, string architecture, bool apply) =>
        Run(root, id, architecture, apply, RuntimeCache.EmbeddedManifest());

    internal static RuntimeCleanupResult Run(string root, string id, string architecture, bool apply, IEnumerable<string> manifest)
    {
        root = PackageRules.SafePath(root);
        var cache = PlatformHandoff.Cache(root, id);
        var expected = RuntimeCache.Parse(manifest, architecture);
        using var installationGate = Exclusive(root, "platform-update.lock");
        if (PlatformHandoff.ReadId(root) == id) throw new IOException("The pending update must be retained.");
        using var jobGate = Exclusive(cache, "apply.lock");
        using var stageGate = Exclusive(cache, "runtime-stage.lock");
        using var useGate = Exclusive(cache, "runtime-use.lock");
        var job = JsonSerializer.Deserialize<PendingPackage>(Read(cache, "package-update.json"))
            ?? throw new IOException("Missing package job.");
        if (job.SchemaVersion != 1 || !PackageRules.SafePath(job.Root).Equals(root, StringComparison.OrdinalIgnoreCase))
            throw new IOException("Package job belongs to another installation.");
        var cancelled = File.Exists(PackageRules.SafePath(Path.Combine(cache, "cancelled.json")));
        if (cancelled)
        {
            if (File.Exists(PackageRules.SafePath(Path.Combine(cache, "apply-started.json"))))
                throw new IOException("An attempted installation must be retained for recovery.");
        }
        else
        {
            using var status = JsonDocument.Parse(Read(cache, "status.json"));
            if (!status.RootElement.TryGetProperty("status", out var value) || value.GetString() is not ("installed" or "restored"))
                throw new IOException("Only completed or cancelled jobs can release their runtime.");
        }
        var directory = PackageRules.SafePath(Path.Combine(cache, "runtime", architecture));
        PackageRules.RequireOutside(root, directory);
        PackageRules.RequireOutside(directory, System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory());
        PackageRules.RequireOutside(directory, Environment.ProcessPath!);
        if (!Directory.Exists(directory)) return new("already-absent", architecture, 0, 0);
        var files = new List<string>();
        var directories = new List<string>();
        Inventory(directory, files, directories);
        long bytes = 0;
        var leases = new List<FileStream>();
        try
        {
            foreach (var file in files)
            {
                var relative = Path.GetRelativePath(directory, file).Replace('\\', '/');
                if (!expected.TryGetValue(relative, out var hash)) throw new IOException("Unrecognized runtime file; retain for inspection.");
                // Check every file and lock before deleting anything. Missing approved files
                // are allowed so interrupted cleanup can resume safely.
                var stream = new FileStream(PackageRules.SafePath(file), FileMode.Open, FileAccess.Read, FileShare.None);
                leases.Add(stream);
                if (!Convert.ToHexString(SHA256.HashData(stream)).Equals(hash, StringComparison.OrdinalIgnoreCase))
                    throw new IOException("Changed runtime file; retain for inspection.");
                bytes += stream.Length;
            }
        }
        finally { foreach (var lease in leases) lease.Dispose(); }
        if (!apply) return new("reviewed", architecture, files.Count, bytes);
        foreach (var file in files) File.Delete(PackageRules.SafePath(file));
        // All names were inventoried beneath the exact derived runtime path; never recurse on deletion.
        foreach (var child in directories.OrderByDescending(p => p.Length)) Directory.Delete(PackageRules.SafePath(child));
        Directory.Delete(PackageRules.SafePath(directory));
        PackageJob.Write(cache, "runtime-cleanup-" + architecture + ".json", new { status = "removed", files = files.Count, bytes });
        return new("removed", architecture, files.Count, bytes);
    }

    private static FileStream Exclusive(string directory, string name) =>
        new(PackageRules.SafePath(Path.Combine(directory, name)), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);

    private static string Read(string directory, string name)
    {
        var path = PackageRules.SafePath(Path.Combine(directory, name));
        if (new FileInfo(path).Length > 8192) throw new IOException("Invalid job metadata size.");
        return File.ReadAllText(path);
    }

    private static void Inventory(string directory, List<string> files, List<string> directories)
    {
        foreach (var entry in Directory.EnumerateFileSystemEntries(directory))
        {
            var path = PackageRules.SafePath(entry);
            if (Directory.Exists(path))
            {
                directories.Add(path);
                Inventory(path, files, directories);
            }
            else files.Add(path);
        }
    }
}
