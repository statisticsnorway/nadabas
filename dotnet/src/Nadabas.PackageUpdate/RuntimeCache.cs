using System.Security.Cryptography;

namespace Nadabas.PackageUpdate;

// Hashes come from the signed updater's embedded resource, never a job-side file.
internal sealed class RuntimeCache : IDisposable
{
    private readonly List<FileStream> leases = [];
    internal string DirectoryPath { get; private init; } = "";
    internal string CoreDirectory { get; private init; } = "";

    internal static string[] EmbeddedManifest()
    {
        using var resource = typeof(RuntimeCache).Assembly.GetManifestResourceStream("Nadabas.RuntimeManifest")
            ?? throw new IOException("Private runtime manifest is missing.");
        using var reader = new StreamReader(resource);
        return reader.ReadToEnd().Split('\n');
    }

    internal static RuntimeCache Open(string cache, string architecture) => Open(cache, architecture, EmbeddedManifest());

    internal static RuntimeCache Open(string cache, string architecture, IEnumerable<string> manifest)
        => Open(cache, architecture, Parse(manifest, architecture));

    private static RuntimeCache Open(string cache, string architecture, Dictionary<string, string> expected)
    {
        var use = Use(cache);
        try
        {
            var result = Verify(Path.Combine(PackageRules.SafePath(cache), "runtime", architecture), expected);
            result.leases.Add(use);
            return result;
        }
        catch { use.Dispose(); throw; }
    }

    private static FileStream Use(string cache) => new(PackageRules.SafePath(Path.Combine(cache, "runtime-use.lock")),
        FileMode.OpenOrCreate, FileAccess.Read, FileShare.Read);

    internal static RuntimeCache? LeaseForExecution(string installation, string cache) =>
        LeaseForExecution(installation, cache, System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory(),
            System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture, EmbeddedManifest());

    internal static RuntimeCache? LeaseForExecution(string installation, string cache, string loadedDirectory,
        System.Runtime.InteropServices.Architecture architecture, IEnumerable<string> manifest)
    {
        cache = PackageRules.SafePath(cache);
        var loaded = PackageRules.SafePath(loadedDirectory);
        PackageRules.RequireOutside(installation, cache);
        PackageRules.RequireOutside(installation, loaded);
        // A cached runtime cannot silently become a shared-runtime job after deletion.
        if (!Directory.Exists(PackageRules.SafePath(Path.Combine(cache, "runtime"))) &&
            !loaded.StartsWith(cache + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) return null;
        var arch = architecture switch
        {
            System.Runtime.InteropServices.Architecture.X86 => "x86",
            System.Runtime.InteropServices.Architecture.X64 => "x64",
            _ => throw new IOException("Unsupported runtime architecture.")
        };
        var lease = Open(cache, arch, manifest);
        if (string.Equals(loaded, lease.CoreDirectory, StringComparison.OrdinalIgnoreCase)) return lease;
        lease.Dispose();
        throw new IOException("The helper did not load the verified cached runtime.");
    }

    internal static RuntimeCache? StageForCurrentRuntime(string installation, string cache)
    {
        var root = PackageRules.SafePath(installation);
        var loaded = PackageRules.SafePath(System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory());
        if (!loaded.Equals(root, StringComparison.OrdinalIgnoreCase) &&
            !loaded.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) return null;
        var architecture = System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture switch
        {
            System.Runtime.InteropServices.Architecture.X86 => "x86",
            System.Runtime.InteropServices.Architecture.X64 => "x64",
            _ => throw new IOException("Unsupported runtime architecture.")
        };
        var expected = Path.Combine(root, "runtime", architecture, "shared", "Microsoft.NETCore.App");
        if (!string.Equals(Path.GetDirectoryName(loaded), expected, StringComparison.OrdinalIgnoreCase))
            throw new IOException("Loaded private runtime has an unexpected location.");
        return Stage(root, cache, architecture);
    }

    internal static RuntimeCache Stage(string installation, string cache, string architecture)
    {
        return Stage(installation, cache, architecture, EmbeddedManifest());
    }

    internal static RuntimeCache Stage(string installation, string cache, string architecture,
        IEnumerable<string> manifest, CancellationToken cancellation = default)
    {
        var expected = Parse(manifest, architecture);
        installation = PackageRules.SafePath(installation);
        cache = PackageRules.SafePath(cache);
        PackageRules.RequireOutside(installation, cache);
        if (!Directory.Exists(cache)) throw new IOException("The package job directory must already exist.");
        using var use = Use(cache);
        using var gate = new FileStream(PackageRules.SafePath(Path.Combine(cache, "runtime-stage.lock")),
            FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
        cancellation.ThrowIfCancellationRequested();
        var destination = PackageRules.SafePath(Path.Combine(cache, "runtime", architecture));
        // Retries may reuse an exact verified copy, but never repair files in place.
        if (Directory.Exists(destination)) return Open(cache, architecture, expected);
        using var source = Verify(Path.Combine(installation, "runtime", architecture), expected);
        var temporary = PackageRules.SafePath(Path.Combine(cache, "runtime-pending-" + Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(temporary);
        foreach (var item in expected)
        {
            cancellation.ThrowIfCancellationRequested();
            var target = PackageRules.SafePath(Path.Combine(temporary, item.Key));
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            using var input = File.OpenRead(PackageRules.SafePath(Path.Combine(source.DirectoryPath, item.Key)));
            using var output = new FileStream(target, FileMode.CreateNew, FileAccess.Write, FileShare.None);
            input.CopyTo(output);
            output.Flush(flushToDisk: true);
        }
        using (Verify(temporary, expected)) { }
        cancellation.ThrowIfCancellationRequested();
        Directory.CreateDirectory(PackageRules.SafePath(Path.GetDirectoryName(destination)!));
        Directory.Move(temporary, destination);
        // Retain read leases on the final names; incomplete copies remain inert for diagnosis.
        return Open(cache, architecture, expected);
    }

    internal static Dictionary<string, string> Parse(IEnumerable<string> manifest, string architecture)
    {
        if (architecture is not ("x86" or "x64")) throw new IOException("Unsupported runtime architecture.");
        var expected = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var raw in manifest)
        {
            var line = raw.TrimEnd('\r');
            if (string.IsNullOrWhiteSpace(line) || line.StartsWith('#')) continue;
            var fields = line.Split('\t');
            if (fields.Length != 3 || fields[0] is not ("x86" or "x64")) throw new IOException("Invalid runtime manifest.");
            if (fields[0] != architecture) continue;
            var relative = fields[1];
            if (relative.Split('/').Any(p => p.Length == 0 || p is "." or ".." || p.EndsWith('.') || p.EndsWith(' ')) ||
                relative.IndexOfAny(['\\', ':']) >= 0 || Path.IsPathRooted(relative) ||
                (relative != "dotnet.exe" && !relative.StartsWith("host/fxr/", StringComparison.Ordinal) &&
                 !relative.StartsWith("shared/Microsoft.NETCore.App/", StringComparison.Ordinal) &&
                 !relative.StartsWith("shared/Microsoft.WindowsDesktop.App/", StringComparison.Ordinal)))
                throw new IOException("Runtime manifest path is outside its payload.");
            if (fields[2].Length != 64 || fields[2].Any(c => !Uri.IsHexDigit(c)) || !expected.TryAdd(relative, fields[2]))
                throw new IOException("Invalid or duplicate runtime hash.");
        }
        var hosts = expected.Keys.Where(p => p.StartsWith("host/fxr/", StringComparison.Ordinal) && p.EndsWith("/hostfxr.dll", StringComparison.Ordinal)).ToArray();
        if (hosts.Length != 1) throw new IOException("Exactly one runtime host must be pinned.");
        var parts = hosts[0].Split('/');
        if (parts.Length != 4 || !Version.TryParse(parts[2], out var version) ||
            version.Major != 10 || version.Minor != 0 || version.Build < 0 || version.Revision != -1 ||
            !expected.ContainsKey("dotnet.exe") ||
            !expected.ContainsKey($"shared/Microsoft.NETCore.App/{version}/coreclr.dll") ||
            !expected.ContainsKey($"shared/Microsoft.WindowsDesktop.App/{version}/System.Windows.Forms.dll"))
            throw new IOException("A matching .NET 10 Desktop Runtime must be pinned.");
        foreach (var path in expected.Keys.Where(p => p != "dotnet.exe"))
            if (path.Split('/')[2] != version.ToString()) throw new IOException("Runtime versions must match.");
        return expected;
    }

    private static RuntimeCache Verify(string directory, Dictionary<string, string> expected)
    {
        var version = expected.Keys.Single(p => p.StartsWith("host/fxr/", StringComparison.Ordinal) && p.EndsWith("/hostfxr.dll", StringComparison.Ordinal)).Split('/')[2];
        var result = new RuntimeCache
        {
            DirectoryPath = PackageRules.SafePath(directory),
            CoreDirectory = PackageRules.SafePath(Path.Combine(directory, "shared", "Microsoft.NETCore.App", version))
        };
        try
        {
            var actual = Files(result.DirectoryPath).Select(p => Path.GetRelativePath(result.DirectoryPath, p).Replace('\\', '/')).ToArray();
            if (actual.Length != expected.Count || actual.Any(p => !expected.ContainsKey(p)))
                throw new IOException("Runtime contains missing or unexpected files.");
            foreach (var item in expected)
            {
                var stream = File.OpenRead(PackageRules.SafePath(Path.Combine(result.DirectoryPath, item.Key)));
                result.leases.Add(stream);
                if (!Convert.ToHexString(SHA256.HashData(stream)).Equals(item.Value, StringComparison.OrdinalIgnoreCase))
                    throw new IOException("Runtime hash mismatch.");
            }
            return result;
        }
        catch { result.Dispose(); throw; }
    }

    private static IEnumerable<string> Files(string directory)
    {
        foreach (var path in Directory.EnumerateFileSystemEntries(directory))
        {
            _ = PackageRules.SafePath(path);
            if (Directory.Exists(path))
            {
                foreach (var file in Files(path)) yield return file;
            }
            else yield return path;
        }
    }

    public void Dispose()
    {
        foreach (var stream in leases) stream.Dispose();
        leases.Clear();
    }
}
