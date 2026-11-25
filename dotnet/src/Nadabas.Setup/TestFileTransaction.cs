using System.Security.Cryptography;
using System.Text.Json;

namespace Nadabas.Setup;

internal sealed record TestPlan(string SourceHash, string? TargetHash);
internal sealed record TestJournal(int SchemaVersion, string State, string SourceHash, string? TargetHash);
internal sealed record TransactionResult(string Status, string? Error = null);

// Deliberately internal: stage 3 operates only inside newly created rehearsal
// directories. This is not the production installer or the executor for plan-install.
internal sealed class TestFileTransaction
{
    internal const string Marker = "NADABAS disposable installation rehearsal v1";
    private readonly string root;
    internal string Source => Path.Combine(root, "incoming.xlam");
    internal string Target => Path.Combine(root, "NADABAS.XLAM");
    internal string Backup => Path.Combine(root, "backup.xlam");
    private string Stage => Path.Combine(root, "staged.xlam");
    private string Journal => Path.Combine(root, "transaction.json");

    internal TestFileTransaction(string directory) => root = Path.GetFullPath(directory);

    internal static TestFileTransaction Create(string parent)
    {
        CheckPath(parent);
        var root = Path.Combine(parent, "nadabas-rehearsal-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        File.WriteAllText(Path.Combine(root, ".nadabas-test-workspace"), Marker);
        return new(root);
    }

    internal string DirectoryPath => root;

    internal TestPlan Plan()
    {
        Validate();
        return new(Hash(Source) ?? throw new IOException("Source missing."), Hash(Target));
    }

    internal TransactionResult Apply(TestPlan plan, Action<string>? checkpoint = null,
        Action? beforeReplace = null)
    {
        Validate();
        using var transactionLock = new FileStream(Path.Combine(root, "transaction.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
        var prepared = false;
        try
        {
            if (File.Exists(Journal)) throw new IOException("Workspace already has a transaction. Recover or use a new workspace.");
            EnsureSnapshot(plan);
            if (plan.SourceHash == plan.TargetHash) return new("unchanged");

            // Open the existing target exclusively before doing any replacement work.
            if (plan.TargetHash is not null)
            {
                if ((File.GetAttributes(Target) & FileAttributes.ReadOnly) != 0)
                    throw new UnauthorizedAccessException("Target is read-only; its attributes will not be changed.");
                using var probe = new FileStream(Target, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            CopyVerified(Source, Stage, plan.SourceHash);
            checkpoint?.Invoke("staged");
            if (plan.TargetHash is not null) CopyVerified(Target, Backup, plan.TargetHash);
            checkpoint?.Invoke("backed-up");
            EnsureSnapshot(plan);
            if (Hash(Stage) != plan.SourceHash) throw new IOException("Staged file changed before replacement.");
            if (plan.TargetHash is not null && Hash(Backup) != plan.TargetHash) throw new IOException("Backup verification failed.");
            WriteJournal(new(1, "prepared", plan.SourceHash, plan.TargetHash));
            prepared = true;
            checkpoint?.Invoke("prepared");
            Validate();
            EnsureSnapshot(plan);
            beforeReplace?.Invoke();
            // Same-directory rename; never delete the target before the replacement is ready.
            File.Move(Stage, Target, overwrite: plan.TargetHash is not null);
            checkpoint?.Invoke("replaced");
            if (Hash(Target) != plan.SourceHash) throw new IOException("Installed file verification failed.");
            WriteJournal(new(1, "installed", plan.SourceHash, plan.TargetHash));
            return new("installed");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            if (!prepared) return new("blocked", ex.Message);
            try { return RecoverCore() with { Error = ex.Message }; }
            catch (Exception recoveryError) when (recoveryError is IOException or UnauthorizedAccessException or JsonException)
            {
                return new("recovery-required", recoveryError.Message);
            }
        }
    }

    internal TransactionResult Recover()
    {
        Validate();
        using var transactionLock = new FileStream(Path.Combine(root, "transaction.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
        return RecoverCore();
    }

    private TransactionResult RecoverCore()
    {
        Validate();
        var journal = JsonSerializer.Deserialize<TestJournal>(File.ReadAllText(Journal))
            ?? throw new IOException("Missing journal data.");
        if (journal.SchemaVersion != 1 || journal.State is not ("prepared" or "installed" or "rolled-back") ||
            !IsHash(journal.SourceHash) || (journal.TargetHash is not null && !IsHash(journal.TargetHash)))
            throw new IOException("Invalid recovery journal.");
        var current = Hash(Target);
        if (journal.State == "rolled-back" && current != journal.TargetHash)
            throw new IOException("Previously restored target changed. Do not overwrite later changes.");
        if (current != journal.SourceHash && current != journal.TargetHash)
            throw new IOException("Target changed unexpectedly. Preserve files for manual recovery.");
        if (current != journal.TargetHash)
        {
            if (journal.TargetHash is null)
            {
                // Preserve the new bytes for inspection instead of deleting them.
                File.Move(Target, Path.Combine(root, "rolled-back-new.xlam"));
            }
            else
            {
                if (Hash(Backup) != journal.TargetHash) throw new IOException("Backup does not match the journal.");
                var restore = Path.Combine(root, "restore.xlam");
                CopyVerified(Backup, restore, journal.TargetHash);
                if (Hash(Target) != current) throw new IOException("Target changed during recovery.");
                File.Move(restore, Target, overwrite: true);
            }
        }
        if (Hash(Target) != journal.TargetHash) throw new IOException("Rollback verification failed.");
        WriteJournal(journal with { State = "rolled-back" });
        return new("rolled-back");
    }

    private void EnsureSnapshot(TestPlan plan)
    {
        Validate();
        if (Hash(Source) != plan.SourceHash || Hash(Target) != plan.TargetHash)
            throw new IOException("Source or target changed after planning.");
    }

    private void WriteJournal(TestJournal journal)
    {
        Validate();
        var pending = Path.Combine(root, "journal.pending");
        using (var stream = new FileStream(pending, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            JsonSerializer.Serialize(stream, journal);
            stream.Flush(flushToDisk: true);
        }
        File.Move(pending, Journal, overwrite: true);
    }

    private void Validate()
    {
        var name = Path.GetFileName(root);
        const string prefix = "nadabas-rehearsal-";
        if (!name.StartsWith(prefix, StringComparison.Ordinal) || !Guid.TryParseExact(name[prefix.Length..], "N", out _))
            throw new IOException("Only generated NADABAS rehearsal directories are accepted.");
        CheckPath(root);
        foreach (var path in Directory.EnumerateFileSystemEntries(root)) CheckPath(path);
        if (File.ReadAllText(Path.Combine(root, ".nadabas-test-workspace")) != Marker)
            throw new IOException("Not a NADABAS disposable test workspace.");
    }

    internal static void CheckPath(string path)
    {
        for (string? current = Path.GetFullPath(path); current is not null; current = Path.GetDirectoryName(current))
        {
            try
            {
                if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("Reparse points are not supported in the rehearsal workspace.");
            }
            catch (FileNotFoundException) { }
            catch (DirectoryNotFoundException) { }
        }
    }

    internal static string? Hash(string path)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return Convert.ToHexString(SHA256.HashData(stream));
        }
        catch (FileNotFoundException) { return null; }
    }

    private static void CopyVerified(string source, string destination, string expectedHash)
    {
        var existingHash = Hash(destination);
        if (existingHash is not null)
        {
            if (existingHash != expectedHash) throw new IOException("Existing staging or backup file has unexpected contents.");
            return;
        }
        using (var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (var output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        {
            input.CopyTo(output);
            output.Flush(flushToDisk: true);
        }
        if (Hash(destination) != expectedHash) throw new IOException("Copied file verification failed.");
    }

    private static bool IsHash(string? value) => value is { Length: 64 } && value.All(Uri.IsHexDigit);
}
