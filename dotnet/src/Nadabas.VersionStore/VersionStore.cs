using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Nadabas.Versioning;

public sealed record ActiveVersion(int SchemaVersion, string Slot);
public sealed record PendingActivation(int SchemaVersion, string PreviousSlot, string TargetSlot);
public sealed record ReadyUpdate(string Slot, string Version);
public sealed record StartupSelection(VersionSelection Selection, string UpdateStatus, string? UpdateError = null);

public sealed record ApprovedRelease(string Version, string Sha256);

public sealed class VersionSelection(string slot, string path, string version, FileStream lease) : IDisposable
{
    public string Slot { get; } = slot;
    public string AddinPath { get; } = path;
    public string ReleaseVersion { get; } = version;
    // A Windows reader lease denies writes/deletion for this session. No pruning
    // is implemented: all old versions are retained even after leases close.
    public void Dispose() => lease.Dispose();
}

public sealed class VersionStore
{
    public const string BaselineHash = "2F7432B9B04C414337EB1AFEB07E2B93F05E4A5A6D151073FA967A2F7EA89A85";
    private const string Marker = "NADABAS version-store prototype v1";
    private readonly string root;
    private readonly IReadOnlyDictionary<string, ApprovedRelease> releases;
    public VersionStore(string path) : this(path, ApprovedReleases.Catalog) { }
    internal VersionStore(string path, IReadOnlyDictionary<string, ApprovedRelease> catalog)
    {
        root = Path.GetFullPath(path);
        releases = new Dictionary<string, ApprovedRelease>(catalog, StringComparer.Ordinal);
    }

    public static VersionStore Create(string path) => Create(path, ApprovedReleases.Catalog);
    internal static VersionStore Create(string path, string hash) => Create(path,
        new Dictionary<string, ApprovedRelease> { ["slot-a"] = new("6.02.005", hash), ["slot-b"] = new("6.02.005", hash) });
    internal static VersionStore Create(string path, IReadOnlyDictionary<string, ApprovedRelease> catalog)
    {
        CheckPath(path);
        if (Directory.Exists(path) || File.Exists(path)) throw new IOException("Use a new prototype directory.");
        Directory.CreateDirectory(path);
        File.WriteAllText(Path.Combine(path, ".nadabas-version-store"), Marker);
        Directory.CreateDirectory(Path.Combine(path, "versions"));
        return new(path, catalog);
    }

    public void Stage(string slot, string source)
    {
        Validate();
        CheckPath(source);
        using var gate = Lock();
        var directory = SlotDirectory(slot);
        var release = Release(slot);
        if (Directory.Exists(directory)) throw new IOException("Version slots are immutable; use a new slot.");
        if (Hash(source) != release.Sha256) throw new IOException("Source does not match the approved release for this slot.");
        var pending = Path.Combine(root, "versions", ".staging-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(pending);
        var file = Path.Combine(pending, AddinName(slot));
        File.Copy(source, file);
        if (Hash(file) != release.Sha256) throw new IOException("Staged bytes do not match the approved release.");
        File.SetAttributes(file, FileAttributes.ReadOnly);
        Directory.Move(pending, directory);
    }

    public void Activate(string slot)
    {
        Validate();
        using var gate = Lock();
        using var validated = SelectSlot(slot);
        WriteRecord("active.json", new ActiveVersion(1, slot));
        File.Delete(Path.Combine(root, "pending-update.json"));
    }

    private void WriteRecord<T>(string name, T record)
    {
        var pending = Path.Combine(root, "record-" + Guid.NewGuid().ToString("N") + ".pending");
        using (var stream = new FileStream(pending, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        {
            JsonSerializer.Serialize(stream, record);
            stream.Flush(true);
        }
        File.Move(pending, Path.Combine(root, name), overwrite: true);
    }

    public PendingActivation QueueActivation(string slot)
    {
        Validate();
        using var gate = Lock();
        using var current = SelectCore();
        using var target = SelectSlot(slot);
        if (Version.Parse(target.ReleaseVersion) <= Version.Parse(current.ReleaseVersion))
            throw new IOException("Only a newer approved release can be queued. Use explicit activation for rollback.");
        var pending = new PendingActivation(1, current.Slot, target.Slot);
        WriteRecord("pending-update.json", pending);
        return pending;
    }

    public void CancelPendingActivation()
    {
        Validate();
        using var gate = Lock();
        File.Delete(Path.Combine(root, "pending-update.json"));
    }

    // Readiness is verified independently of the startup path. Never activate
    // merely because the UI asks for status, or trust version text in a sidecar.
    public ReadyUpdate? GetReadyUpdate()
    {
        Validate();
        using var gate = Lock();
        var path = Path.Combine(root, "pending-update.json");
        if (!File.Exists(path)) return null;
        if (new FileInfo(path).Length > 1024) throw new IOException("Pending update record is too large.");
        var pending = JsonSerializer.Deserialize<PendingActivation>(File.ReadAllText(path));
        if (pending is null || pending.SchemaVersion != 1) throw new IOException("Invalid pending update record.");
        using var current = SelectCore();
        if (current.Slot == pending.TargetSlot) return null;
        if (current.Slot != pending.PreviousSlot) throw new IOException("Active selection changed after queueing.");
        using var target = SelectSlot(pending.TargetSlot);
        if (Version.Parse(target.ReleaseVersion) <= Version.Parse(current.ReleaseVersion))
            throw new IOException("Pending update is not a newer release.");
        return new(target.Slot, target.ReleaseVersion);
    }

    public StartupSelection SelectForStartup()
    {
        Validate();
        using var gate = Lock();
        var current = SelectCore();
        var pendingPath = Path.Combine(root, "pending-update.json");
        if (!File.Exists(pendingPath)) return new(current, "none");
        VersionSelection? target = null;
        try
        {
            if (new FileInfo(pendingPath).Length > 1024) throw new IOException("Pending update record is too large.");
            var pending = JsonSerializer.Deserialize<PendingActivation>(File.ReadAllText(pendingPath));
            if (pending is null || pending.SchemaVersion != 1) throw new IOException("Invalid pending update record.");
            // A crash may occur after the atomic pointer switch but before cleanup.
            if (current.Slot == pending.TargetSlot)
            {
                File.Delete(pendingPath);
                return new(current, "already-applied");
            }
            if (current.Slot != pending.PreviousSlot) throw new IOException("Active selection changed after the update was queued.");
            target = SelectSlot(pending.TargetSlot);
            if (Version.Parse(target.ReleaseVersion) <= Version.Parse(current.ReleaseVersion))
                throw new IOException("Pending automatic update is not an upgrade.");
            WriteRecord("active.json", new ActiveVersion(1, target.Slot));
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException or ArgumentException)
        {
            target?.Dispose();
            return new(current, "rejected", ex.Message);
        }
        current.Dispose();
        try { File.Delete(pendingPath); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return new(target!, "applied-cleanup-pending", ex.Message);
        }
        return new(target!, "applied");
    }

    public VersionSelection Select()
    {
        Validate();
        using var gate = Lock();
        return SelectCore();
    }

    private VersionSelection SelectCore()
    {
        var pointer = Path.Combine(root, "active.json");
        CheckPath(pointer);
        if (new FileInfo(pointer).Length > 1024) throw new IOException("Invalid active-version record size.");
        var active = JsonSerializer.Deserialize<ActiveVersion>(File.ReadAllText(pointer));
        if (active is null || active.SchemaVersion != 1) throw new IOException("Invalid active-version record.");
        return SelectSlot(active.Slot);
    }

    private VersionSelection SelectSlot(string slot)
    {
        var path = Path.Combine(SlotDirectory(slot), AddinName(slot));
        var release = Release(slot);
        CheckPath(path);
        var lease = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        try
        {
            if (Convert.ToHexString(SHA256.HashData(lease)) != release.Sha256) throw new IOException("Selected version differs from its approved release.");
            return new(slot, path, release.Version, lease);
        }
        catch { lease.Dispose(); throw; }
    }

    private ApprovedRelease Release(string slot) => releases.TryGetValue(slot, out var release)
        ? release : throw new IOException("This slot is not in the loader's approved release catalog.");

    private string SlotDirectory(string slot)
    {
        if (slot is null || !Regex.IsMatch(slot, @"^slot-[a-z0-9]{1,16}$")) throw new IOException("Invalid prototype slot name.");
        var path = Path.Combine(root, "versions", slot);
        CheckPath(path);
        return path;
    }
    private string AddinName(string slot) => "NADABAS.TEST." + slot + ".xlam";
    private FileStream Lock() => new(Path.Combine(root, "store.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    private void Validate()
    {
        CheckPath(root);
        foreach (var name in new[] { ".nadabas-version-store", "store.lock", "versions", "active.json", "pending-update.json" }) CheckPath(Path.Combine(root, name));
        if (File.ReadAllText(Path.Combine(root, ".nadabas-version-store")) != Marker) throw new IOException("Not a prototype version store.");
    }
    private static string Hash(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }
    private static void CheckPath(string path)
    {
        for (string? current = Path.GetFullPath(path); current is not null; current = Path.GetDirectoryName(current))
        {
            try
            {
                if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0) throw new IOException("Prototype store does not support linked paths.");
            }
            catch (FileNotFoundException) { }
            catch (DirectoryNotFoundException) { }
        }
    }
}
