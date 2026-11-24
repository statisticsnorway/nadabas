using System.Reflection.PortableExecutable;
using System.Security;

namespace Nadabas.Setup;

public sealed record ExcelRegistration(string Source, string? Path);
public sealed record ExcelCandidate(string Path, string Status, string? Architecture,
    IReadOnlyList<string> Sources, string Evidence);
public sealed record ExcelDiscoveryResult(string Status, ExcelCandidate? Selected,
    IReadOnlyList<ExcelCandidate> Candidates);

// Read-only App Paths inspection. Selection is evidence, not permission to install.
public static class ExcelDiscovery
{
    public static ExcelDiscoveryResult Inspect(IEnumerable<ExcelRegistration> registrations)
    {
        var paths = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        var candidates = new List<ExcelCandidate>();
        foreach (var registration in registrations)
        {
            if (string.IsNullOrWhiteSpace(registration.Path)) continue;
            var path = registration.Path.Trim();
            if (path.StartsWith('"') && path.EndsWith('"') && path.Length > 1) path = path[1..^1];
            if (!Path.IsPathFullyQualified(path) ||
                !string.Equals(Path.GetFileName(path), "EXCEL.EXE", StringComparison.OrdinalIgnoreCase))
            {
                candidates.Add(new(path, "invalid-path", null, [registration.Source],
                    "Expected an absolute Excel executable path, without command arguments."));
                continue;
            }
            try { path = Path.GetFullPath(path); }
            catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
            {
                candidates.Add(new(path, "invalid-path", null, [registration.Source], "Invalid executable path."));
                continue;
            }
            if (!paths.TryGetValue(path, out var sources)) paths.Add(path, sources = []);
            if (!sources.Contains(registration.Source, StringComparer.Ordinal)) sources.Add(registration.Source);
        }
        foreach (var entry in paths.OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase))
            candidates.Add(InspectFile(entry.Key, entry.Value));

        // Stale or unreadable competing registrations must not silently select another install.
        if (candidates.Count == 0) return new("not-found", null, candidates);
        if (candidates.Count > 1) return new("ambiguous", null, candidates);
        var candidate = candidates[0];
        return new(candidate.Status, candidate.Status == "verified" ? candidate : null, candidates);
    }

    private static ExcelCandidate InspectFile(string path, IReadOnlyList<string> sources)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var reader = new PEReader(stream);
            var headers = reader.PEHeaders;
            var flags = headers.CoffHeader.Characteristics;
            if (headers.PEHeader is null || headers.CoffHeader.NumberOfSections == 0 ||
                (flags & Characteristics.ExecutableImage) == 0 || (flags & Characteristics.Dll) != 0)
                return new(path, "invalid-image", null, sources, "Not a PE executable image.");
            var architecture = headers.CoffHeader.Machine switch
            {
                Machine.I386 when headers.PEHeader.Magic == PEMagic.PE32 => "x86",
                Machine.Amd64 when headers.PEHeader.Magic == PEMagic.PE32Plus => "x64",
                _ => null
            };
            return architecture is null
                ? new(path, "unsupported-image", null, sources, "Unsupported machine type or inconsistent PE header.")
                : new(path, "verified", architecture, sources,
                    "Architecture verified from PE headers only; publisher, Office version, OS compatibility and startup are not verified.");
        }
        catch (Exception ex) when (ex is FileNotFoundException or DirectoryNotFoundException)
        { return new(path, "missing", null, sources, "Registered executable does not exist."); }
        catch (BadImageFormatException)
        { return new(path, "invalid-image", null, sources, "Malformed or truncated PE image."); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or SecurityException or ArgumentException or NotSupportedException)
        { return new(path, "error", null, sources, "Executable inspection failed: " + ex.GetType().Name); }
    }
}
