using System.Security;
using System.Security.Cryptography;

namespace Nadabas.Setup;

public sealed record FileSnapshot(string Path, long Length, string Sha256);
public sealed record PlannedStep(string Action, string? Source, string Destination);
public sealed record InstallationPlan(int SchemaVersion, string Kind, string Status, bool CanApply,
    string SourcePath, string TargetPath, FileSnapshot? Source, FileSnapshot? ExistingTarget,
    IReadOnlyList<PlannedStep> Steps, IReadOnlyList<string> Blockers,
    IReadOnlyList<string> RequiredChecks);

public static class InstallationPlanner
{
    public static InstallationPlan Create(string sourcePath, string targetDirectory)
    {
        sourcePath = Path.GetFullPath(sourcePath);
        targetDirectory = Path.GetFullPath(targetDirectory);
        var targetPath = Path.Combine(targetDirectory, "NADABAS.XLAM");
        var blockers = new List<string>();
        var steps = new List<PlannedStep>();
        FileSnapshot? source = null;
        FileSnapshot? existing = null;
        try
        {
            if (!Diagnostics.IsNadabasAddin(Path.GetFileName(sourcePath)))
                blockers.Add("Source must be named NADABAS.xlam or NADABAS.<version>.xlam.");
            RejectLinks(sourcePath, blockers);
            RejectLinks(targetDirectory, blockers);
            if (blockers.Count == 0)
            {
                source = Snapshot(sourcePath);
                if (source.Length == 0)
                    blockers.Add("Source add-in is empty.");

                // Do not use Directory.Exists: access errors must not look like a new installation.
                string[] files;
                try
                {
                    files = Directory.GetFiles(targetDirectory);
                }
                catch (DirectoryNotFoundException)
                {
                    files = [];
                }

                var candidates = files.Where(f => Diagnostics.IsNadabasAddin(Path.GetFileName(f))).ToArray();
                var canonical = candidates.Where(f => Path.GetFileName(f).Equals("NADABAS.XLAM", StringComparison.OrdinalIgnoreCase)).ToArray();
                if (canonical.Length == 1)
                    targetPath = canonical[0]; // Preserve actual spelling on case-sensitive file systems.
                if (candidates.Length > canonical.Length || canonical.Length > 1)
                    blockers.Add("Other or ambiguous NADABAS add-ins exist in the target directory. Review them in Excel; no removal is planned.");
                RejectLinks(targetPath, blockers);
                if (blockers.Count == 0 && canonical.Length == 1)
                    existing = Snapshot(targetPath);
                if (blockers.Count == 0)
                {
                    if (existing?.Sha256 == source.Sha256)
                        steps.Add(new("keep-identical-file", null, targetPath));
                    else
                    {
                        if (existing is not null)
                            steps.Add(new("backup-required-before-replacement", targetPath,
                                Path.Combine(targetDirectory, ".nadabas-backups", existing.Sha256, "NADABAS.XLAM")));
                        steps.Add(new(existing is null ? "install-file" : "replace-file", sourcePath, targetPath));
                    }
                }
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or SecurityException)
        {
            blockers.Add("File inspection failed: " + ex.GetType().Name + ". No installation steps can be proposed.");
        }

        if (blockers.Count > 0)
            steps.Clear();
        return new(1, "installation-plan", blockers.Count == 0 ? "review-required" : "blocked", false,
            sourcePath, targetPath, source, existing, steps, blockers,
            [
                "Verify the approved release, XLAM contents/version and Office VBA publisher signature. SHA-256 proves byte identity only.",
                "Confirm the target is Excel's intended add-in directory and inspect the registered add-in paths.",
                "Before applying, close Excel and recheck source and target hashes, file locks, path links and write permissions.",
                "Implement and test backup verification, staged replacement and rollback. A proposed backup path is not an existing backup.",
                "Activate/register the add-in through a tested Excel workflow, then verify startup, version and database operations."
            ]);
    }

    private static FileSnapshot Snapshot(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        var length = stream.Length;
        return new(path, length, Convert.ToHexString(SHA256.HashData(stream)));
    }

    private static void RejectLinks(string path, List<string> blockers)
    {
        // Check existing ancestors too, including junctions. No files are followed for hashing
        // when a link is observed. This is a snapshot, not protection against later path changes.
        for (string? current = path; current is not null; current = Path.GetDirectoryName(current))
        {
            try
            {
                var attributes = File.GetAttributes(current);
                if (current != path && (attributes & FileAttributes.Directory) == 0)
                {
                    blockers.Add("A source or target ancestor is a file, not a directory.");
                    return;
                }
                if ((attributes & FileAttributes.ReparsePoint) != 0)
                {
                    blockers.Add("A source or target path contains a symbolic link, junction or reparse point. Use a regular local path for this prototype.");
                    return;
                }
                if (current == path && (attributes & FileAttributes.Directory) != 0 && Diagnostics.IsNadabasAddin(Path.GetFileName(path)))
                {
                    blockers.Add("An add-in file path refers to a directory.");
                    return;
                }
            }
            catch (FileNotFoundException) { }
            catch (DirectoryNotFoundException) { }
        }
    }
}
