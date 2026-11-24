using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using Nadabas.PackageUpdate;

// Engineering probe only. Generated fixture hashes are not release approvals.
if (!OperatingSystem.IsWindows()) throw new IOException("Use Windows for runtime lease acceptance.");
var architecture = RuntimeInformation.ProcessArchitecture.ToString().ToLowerInvariant();
if (args.Length == 4 && args[0] == "child")
{
    using var lease = RuntimeCache.LeaseForExecution(args[1], args[2], RuntimeEnvironment.GetRuntimeDirectory(),
        RuntimeInformation.ProcessArchitecture, File.ReadAllLines(args[3]));
    if (lease is null) throw new IOException("The child must use the private cached runtime.");
    Console.WriteLine("ready:" + architecture);
    if (Console.ReadLine() != "exit") throw new IOException("Missing parent completion signal.");
    return;
}
if (args.Length != 3 || args[0] != "probe") throw new IOException("Use: probe <runtime-source> <report-parent>");
var source = PackageRules.SafePath(args[1]);
var run = PackageRules.SafePath(Path.Combine(args[2], "runtime-handoff-" + Guid.NewGuid().ToString("N")));
PackageRules.RequireOutside(source, run);
Directory.CreateDirectory(run);
var installation = Path.Combine(run, "installation");
var runtime = Path.Combine(installation, "runtime", architecture);
var cache = Path.Combine(run, "job");
Directory.CreateDirectory(cache);
var sourceFiles = PlainFiles(source).ToArray();
var manifest = sourceFiles.Select(path => architecture + "\t" + Path.GetRelativePath(source, path).Replace('\\', '/') + "\t" + Hash(path)).ToArray();
var expected = RuntimeCache.Parse(manifest, architecture);
var manifestPath = Path.Combine(run, "fixture-manifest.txt");
File.WriteAllLines(manifestPath, manifest);
foreach (var file in expected.Keys)
{
    var destination = Path.Combine(runtime, file);
    Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
    File.Copy(Path.Combine(source, file), destination);
}
using var parentLease = RuntimeCache.Stage(installation, cache, architecture, manifest);
var start = new ProcessStartInfo(Environment.ProcessPath!)
{
    UseShellExecute = false,
    CreateNoWindow = true,
    RedirectStandardInput = true,
    RedirectStandardOutput = true,
    RedirectStandardError = true
};
foreach (var value in new[] { "child", installation, cache, manifestPath }) start.ArgumentList.Add(value);
start.Environment["DOTNET_ROOT"] = parentLease.DirectoryPath;
start.Environment["DOTNET_ROOT_" + architecture.ToUpperInvariant()] = parentLease.DirectoryPath;
using var child = Process.Start(start) ?? throw new IOException("Could not start isolated child.");
var errors = child.StandardError.ReadToEndAsync();
try
{
    var ready = await child.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(60));
    if (ready != "ready:" + architecture) throw new IOException("The child did not acquire the expected runtime.");
    parentLease.Dispose();
    var protectedFile = Path.Combine(cache, "runtime", architecture, "dotnet.exe");
    var blocked = false;
    try { using var writable = new FileStream(protectedFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None); }
    catch (IOException) { blocked = true; }
    if (!blocked) throw new IOException("Child did not retain its runtime leases after parent release.");
    await child.StandardInput.WriteLineAsync("exit");
    await child.StandardInput.FlushAsync();
    await child.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(60));
    if (child.ExitCode != 0) throw new IOException("Child failed: " + await errors);
    using (new FileStream(protectedFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
    RemoveFixtureRuntime(run, runtime, expected);
    RemoveFixtureRuntime(run, Path.Combine(cache, "runtime", architecture), expected);
    File.WriteAllText(Path.Combine(run, "report.json"), JsonSerializer.Serialize(new
    {
        passed = true,
        architecture,
        scope = "real-runtime-two-process-lease-handoff",
        files = expected.Count,
        sourceRuntime = source,
        fixtureRuntimeCopiesRemoved = true,
        excelTested = false,
        installerTested = false,
        signedReleaseTested = false
    }, new JsonSerializerOptions { WriteIndented = true }));
    Console.WriteLine(Path.Combine(run, "report.json"));
}
catch
{
    if (!child.HasExited) { child.Kill(); await child.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10)); }
    throw;
}

static string Hash(string path)
{
    using var file = File.OpenRead(PackageRules.SafePath(path));
    return Convert.ToHexString(SHA256.HashData(file));
}

static IEnumerable<string> PlainFiles(string root)
{
    foreach (var entry in Directory.EnumerateFileSystemEntries(PackageRules.SafePath(root)))
    {
        _ = PackageRules.SafePath(entry);
        if (Directory.Exists(entry))
        {
            foreach (var child in PlainFiles(entry)) yield return child;
        }
        else yield return entry;
    }
}

static void RemoveFixtureRuntime(string run, string target, Dictionary<string, string> expected)
{
    target = PackageRules.SafePath(target);
    if (!target.StartsWith(run + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
        throw new IOException("Fixture cleanup escaped its generated run directory.");
    var files = PlainFiles(target).ToArray();
    if (files.Length != expected.Count) throw new IOException("Unexpected fixture inventory.");
    foreach (var file in files)
        if (!expected.TryGetValue(Path.GetRelativePath(target, file).Replace('\\', '/'), out var hash) || Hash(file) != hash)
            throw new IOException("Changed fixture retained for inspection.");
    foreach (var file in files) File.Delete(file);
    foreach (var directory in Directory.GetDirectories(target, "*", SearchOption.AllDirectories).OrderByDescending(p => p.Length))
        Directory.Delete(PackageRules.SafePath(directory));
    Directory.Delete(target);
}
