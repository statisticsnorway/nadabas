using Microsoft.Win32;
using System.Runtime.Versioning;

namespace Nadabas.PackageUpdate;

internal sealed record UpdatePolicy(bool Enabled, bool Automatic, string? PinnedVersion)
{
    internal void RequireAllowed(bool manual, string? version = null)
    {
        if (!Enabled || (!manual && !Automatic)) throw new IOException("Updates are disabled by policy or user preference.");
        if (version is not null && PinnedVersion is not null && PackageRules.Normalize(version) != PackageRules.Normalize(PinnedVersion))
            throw new IOException("This version is not permitted by the IT version pin.");
    }

    [SupportedOSPlatform("windows")]
    internal static UpdatePolicy Read()
    {
        using var machine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
        using var policy = machine.OpenSubKey(@"Software\Policies\NADABAS\Updates");
        using var user = Registry.CurrentUser.OpenSubKey(@"Software\NADABAS\Updates");
        bool Flag(RegistryKey? key, string name, bool fallback)
        {
            var value = key?.GetValue(name);
            return value is null ? fallback : value is int number && number == 1;
        }
        var rawPin = policy?.GetValue("PinnedVersion");
        if (rawPin is not null && rawPin is not string) throw new IOException("Invalid IT version pin.");
        var pin = rawPin as string;
        if (!string.IsNullOrEmpty(pin)) _ = PackageRules.Normalize(pin);
        return new(Flag(policy, "Enabled", true), Flag(policy, "Automatic", true) && Flag(user, "Automatic", true),
            string.IsNullOrEmpty(pin) ? null : pin);
    }
}
