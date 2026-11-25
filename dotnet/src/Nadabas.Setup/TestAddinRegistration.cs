using Microsoft.Win32;
using System.Runtime.Versioning;

namespace Nadabas.Setup;

internal static class TestAddinRegistration
{
    internal static bool RefersToTestFile(string name, object? value, string testPath) =>
        string.Equals(name, testPath, StringComparison.OrdinalIgnoreCase) ||
        (name.StartsWith("OPEN", StringComparison.OrdinalIgnoreCase) && value is string text &&
         new[] { testPath, "\"" + testPath + "\"", "/R \"" + testPath + "\"" }
             .Contains(text, StringComparer.OrdinalIgnoreCase));

    [SupportedOSPlatform("windows")]
    internal static void RemoveTestEntry(string version, string testPath)
    {
        if (!System.Text.RegularExpressions.Regex.IsMatch(version, @"^\d+\.\d+$"))
            throw new IOException("Unexpected Office version; test registration was not removed.");
        foreach (var view in new[] { RegistryView.Registry32, RegistryView.Registry64 })
        {
            using var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
            foreach (var section in new[] { "Add-in Manager", "Options" })
            {
                using var key = root.OpenSubKey($@"Software\Microsoft\Office\{version}\Excel\{section}", writable: true);
                if (key is null) continue;
                foreach (var name in key.GetValueNames())
                    if (RefersToTestFile(name, key.GetValue(name), testPath)) key.DeleteValue(name);
            }
        }
    }
}
