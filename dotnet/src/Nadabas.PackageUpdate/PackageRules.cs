using Microsoft.Win32.SafeHandles;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

namespace Nadabas.PackageUpdate;

internal static class PackageRules
{
    internal static Version Normalize(string value)
    {
        var version = Version.Parse(value);
        return new(version.Major, version.Minor, Math.Max(0, version.Build), Math.Max(0, version.Revision));
    }

    internal static void RequireUpgrade(string current, string incoming)
    {
        if (Normalize(incoming) <= Normalize(current)) throw new IOException("Package must be newer than the installed platform. Use explicit repair or rollback instead.");
    }

    internal static string SafePath(string path)
    {
        if (!Path.IsPathFullyQualified(path)) throw new IOException("An absolute path is required.");
        var full = Path.TrimEndingDirectorySeparator(Path.GetFullPath(path));
        for (var item = full; !string.IsNullOrEmpty(item); item = Path.GetDirectoryName(item))
        {
            if ((File.Exists(item) || Directory.Exists(item)) && (File.GetAttributes(item) & FileAttributes.ReparsePoint) != 0)
                throw new IOException("Filesystem links are not supported.");
        }
        return full;
    }

    internal static void RequireOutside(string installation, string candidate)
    {
        var root = Path.TrimEndingDirectorySeparator(Path.GetFullPath(installation));
        var path = Path.TrimEndingDirectorySeparator(Path.GetFullPath(candidate));
        if (path.Equals(root, StringComparison.OrdinalIgnoreCase) ||
            path.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            throw new IOException("The update cache and running helper must be outside the installation being replaced.");
    }

    internal static void RequireCachedExecutable(string actual, string expected)
    {
        // Excel and Windows may spell the path to an executable differently.
        // Compare the normalized names Windows returns for the two open files.
        using var running = File.OpenRead(SafePath(actual));
        using var cached = File.OpenRead(SafePath(expected));
        var runningName = FinalName(running.SafeFileHandle);
        var cachedName = FinalName(cached.SafeFileHandle);
        if (!runningName.Equals(cachedName, StringComparison.Ordinal))
            throw new IOException($"Restart helper location mismatch. Running: {runningName}. Expected: {cachedName}.");
    }

    private static string FinalName(SafeFileHandle handle)
    {
        var buffer = new StringBuilder(512);
        while (true)
        {
            var length = GetFinalPathNameByHandleW(handle, buffer, (uint)buffer.Capacity, 0);
            if (length == 0) throw new IOException("Cannot resolve restart helper location.", new Win32Exception(Marshal.GetLastWin32Error()));
            if (length < buffer.Capacity) return buffer.ToString();
            buffer.EnsureCapacity(checked((int)length + 1));
        }
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
    private static extern uint GetFinalPathNameByHandleW(SafeFileHandle handle, StringBuilder path, uint size, uint flags);
}
