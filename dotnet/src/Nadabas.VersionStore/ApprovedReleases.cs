namespace Nadabas.Versioning;

// This allowlist is compiled into the signed loader and CLI. A writable sidecar
// manifest must never be able to admit arbitrary VBA content into the loader.
internal static class ApprovedReleases
{
    internal static readonly IReadOnlyDictionary<string, ApprovedRelease> Catalog =
        new Dictionary<string, ApprovedRelease>(StringComparer.Ordinal)
        {
            ["slot-a"] = new("6.02.005", VersionStore.BaselineHash),
            ["slot-b"] = new("6.02.005", VersionStore.BaselineHash),
            // Signed, version-number-only fixture; NOT a published VBA release.
            ["slot-test602006"] = new("6.02.006", "852CC76872428186AC0F8B5C6AA906AE438DABEE6BEF8061E739858DCCD052B8"),
            // Signed isolated fixture with the managed update bridge; NOT published.
            ["slot-test602007"] = new("6.02.007", "F13C5319DF56FA7F7A02B912F98B3B062EA2A4586FB83A86F533A6FEEF01887E"),
            ["slot-test602008"] = new("6.02.008", "C2F70BBA9F93CBE78FA695C6E48057C9092B1FC976C7F788718590B26D92D126"),
            // Real user-supplied beta, distinct from the simulated 006 fixture.
            ["slot-beta602006"] = new("6.02.006", "82F22D19D66A6D682E5A5FBE7C15D7013929D36BF37820FAF915617687D3872D"),
            ["slot-beta602006b"] = new("6.02.006", "BDF847E456764A7D335708471927547B780E33706BEB2A5635337F1FCF3135ED"),
            // Signed multilingual demo; not a published VBA release.
            ["slot-demo602006"] = new("6.02.006", "8D87B0295547DB6F250F99EEC60781ACB1F54F8FF596559253DF55F1628799E0"),
            // Signed issue 47 and Exchange-placement test; not a published release.
            ["slot-fix602009"] = new("6.02.009", "D53D492724997C08596DBEACAE4695B30786B3695310D13035CD61541A829750")
        };
}
