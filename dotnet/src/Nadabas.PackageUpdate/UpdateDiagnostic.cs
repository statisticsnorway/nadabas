using System.Text.Json;

namespace Nadabas.PackageUpdate;

public enum UpdateStage { Download, Verification, Restart }

public sealed record UpdateFailure(string Type, string HResult);

// Deliberately exclude exception messages, file paths, URLs, Data and stack traces:
// those may contain user content or credentials from an external component.
public sealed record UpdateDiagnostic(string Reference, DateTimeOffset TimeUtc, string Stage,
    string ComponentVersion, IReadOnlyList<UpdateFailure> Failures)
{
    public static UpdateDiagnostic Create(UpdateStage stage, Exception error)
    {
        var failures = new List<UpdateFailure>();
        for (Exception? current = error; current is not null && failures.Count < 8; current = current.InnerException)
            failures.Add(new(current.GetType().Name, $"0x{current.HResult:X8}"));
        return new(Guid.NewGuid().ToString("N"), DateTimeOffset.UtcNow, stage.ToString(),
            typeof(UpdateDiagnostic).Assembly.GetName().Version?.ToString() ?? "unknown", failures);
    }

    public string ToJson() => JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });

    public string? TrySave(string root)
    {
        string? temporary = null;
        try
        {
            // Fixed allowlisted filenames bound retention to one report per stage.
            if (!Enum.TryParse<UpdateStage>(Stage, out var stage) || !Enum.IsDefined(stage)) return null;
            var directory = PackageRules.SafePath(Path.Combine(root, "diagnostics"));
            Directory.CreateDirectory(directory);
            var path = PackageRules.SafePath(Path.Combine(directory, $"last-{stage.ToString().ToLowerInvariant()}.json"));
            temporary = PackageRules.SafePath(Path.Combine(directory, Guid.NewGuid().ToString("N") + ".tmp"));
            using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                JsonSerializer.Serialize(file, this);
            File.Move(temporary, path, overwrite: true);
            return path;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            return null; // Diagnostics must never prevent the user from continuing work.
        }
        finally
        {
            if (temporary is not null)
            {
                try { File.Delete(temporary); }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
            }
        }
    }
}
