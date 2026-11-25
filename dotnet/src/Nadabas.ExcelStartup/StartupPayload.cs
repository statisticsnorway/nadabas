using System.Security.Cryptography;

namespace Nadabas.ExcelStartup;

// The manifest is embedded in the startup XLL, never trusted from a writable sidecar.
internal sealed class StartupPayload : IDisposable
{
    private readonly List<FileStream> leases = [];
    internal string Loader { get; private set; } = "";
    internal string Runtime { get; private set; } = "";

    internal static string Architecture(int pointerSize) => pointerSize switch
    {
        4 => "x86",
        8 => "x64",
        _ => throw new IOException("Unsupported Excel process architecture.")
    };

    internal static StartupPayload Verify(string root, string architecture, IEnumerable<string> manifest,
        Action<string> verifyPublisher)
    {
        if (architecture != "x86" && architecture != "x64") throw new IOException("Unsupported architecture.");
        var payload = new StartupPayload();
        try
        {
            root = Path.GetFullPath(root);
            var expected = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in manifest)
            {
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#", StringComparison.Ordinal)) continue;
                var fields = line.Split('\t');
                if (fields.Length != 3 || (fields[0] != "x86" && fields[0] != "x64")) throw new IOException("Invalid startup manifest.");
                if (fields[0] != architecture) continue;
                var relative = fields[1];
                if (relative.Split('/').Any(p => p.Length == 0 || p == "." || p == "..") ||
                    relative.IndexOfAny(['\\', ':']) >= 0 || Path.IsPathRooted(relative)) throw new IOException("Invalid manifest path.");
                if (relative != "Nadabas.Loader." + architecture + ".xll" &&
                    !relative.StartsWith("runtime/" + architecture + "/", StringComparison.Ordinal)) throw new IOException("Manifest path outside selected payload.");
                if (fields[2].Length != 64 || fields[2].Any(c => !Uri.IsHexDigit(c)) || expected.ContainsKey(relative))
                    throw new IOException("Invalid or duplicate manifest hash.");
                expected.Add(relative, fields[2]);
            }
            var loaderName = "Nadabas.Loader." + architecture + ".xll";
            if (!expected.ContainsKey(loaderName) || !expected.ContainsKey("runtime/" + architecture + "/dotnet.exe"))
                throw new IOException("Startup payload is not configured for this Excel architecture.");
            var runtimePrefix = "runtime/" + architecture + "/";
            var hosts = expected.Keys.Where(p => p.StartsWith(runtimePrefix + "host/fxr/", StringComparison.Ordinal)
                && p.EndsWith("/hostfxr.dll", StringComparison.Ordinal)).ToArray();
            if (hosts.Length != 1) throw new IOException("Exactly one private host version is required.");
            var hostParts = hosts[0].Split('/');
            if (hostParts.Length != 6 || !Version.TryParse(hostParts[4], out var version) ||
                version.Major != 10 || version.Minor != 0 || version.Build < 0 || version.Revision != -1 ||
                !expected.ContainsKey(runtimePrefix + "shared/Microsoft.NETCore.App/" + version + "/coreclr.dll") ||
                !expected.ContainsKey(runtimePrefix + "shared/Microsoft.WindowsDesktop.App/" + version + "/System.Windows.Forms.dll"))
                throw new IOException("A complete, matching .NET 10 Desktop Runtime must be pinned.");
            payload.Runtime = Path.Combine(root, "runtime", architecture);
            payload.Loader = Path.Combine(root, loaderName);
            // Reject links before recursive enumeration; no traversal outside the owned tree.
            RequirePlainPath(payload.Runtime);
            var actual = EnumeratePlainFiles(payload.Runtime).Select(p =>
                "runtime/" + architecture + "/" + p.Substring(payload.Runtime.Length + 1).Replace('\\', '/')).ToArray();
            if (actual.Length != expected.Count - 1 || actual.Any(p => !expected.ContainsKey(p)))
                throw new IOException("Private runtime contains missing or unapproved files.");
            foreach (var item in expected)
            {
                var path = Path.Combine(root, item.Key.Replace('/', Path.DirectorySeparatorChar));
                RequirePlainPath(path);
                var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                payload.leases.Add(stream);
                using var sha = SHA256.Create();
                var hash = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "");
                if (!string.Equals(hash, item.Value, StringComparison.OrdinalIgnoreCase)) throw new IOException("Startup payload hash mismatch.");
            }
            verifyPublisher(payload.Loader);
            return payload;
        }
        catch { payload.Dispose(); throw; }
    }

    private static IEnumerable<string> EnumeratePlainFiles(string directory)
    {
        foreach (var path in Directory.EnumerateFileSystemEntries(directory))
        {
            RequirePlainPath(path);
            if (Directory.Exists(path))
            {
                foreach (var file in EnumeratePlainFiles(path)) yield return file;
            }
            else yield return path;
        }
    }

    private static void RequirePlainPath(string path)
    {
        for (var current = path; !string.IsNullOrEmpty(current); current = Path.GetDirectoryName(current))
            if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                throw new IOException("Startup through filesystem links is not supported.");
    }

    internal bool Register(Func<string, bool> register)
    {
        var variables = new[] { "DOTNET_ROOT", "DOTNET_ROOT_X86", "DOTNET_ROOT_X64" };
        var previous = variables.Select(Environment.GetEnvironmentVariable).ToArray();
        try
        {
            foreach (var name in variables) Environment.SetEnvironmentVariable(name, Runtime, EnvironmentVariableTarget.Process);
            return register(Loader);
        }
        finally
        {
            for (var i = 0; i < variables.Length; i++)
                Environment.SetEnvironmentVariable(variables[i], previous[i], EnvironmentVariableTarget.Process);
        }
    }

    public void Dispose()
    {
        foreach (var stream in leases) stream.Dispose();
        leases.Clear();
    }
}
