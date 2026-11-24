using System.Diagnostics;
using System.Reflection.PortableExecutable;
using System.Runtime.InteropServices;

namespace Nadabas.Hosting;

// Shared source: use the runtime actually hosting us, never an inherited DOTNET_ROOT.
internal static class RuntimeChildProcess
{
    internal static string Configure(ProcessStartInfo command, string installation) =>
        Configure(command, installation, RuntimeEnvironment.GetRuntimeDirectory(), RuntimeInformation.ProcessArchitecture);

    internal static string Configure(ProcessStartInfo command, string installation, string runtimeDirectory, Architecture architecture)
    {
        var root = Path.TrimEndingDirectorySeparator(Path.GetFullPath(installation));
        var runtime = new DirectoryInfo(Path.GetFullPath(runtimeDirectory));
        var arch = architecture switch { Architecture.X86 => "x86", Architecture.X64 => "x64", _ => throw new IOException("Unsupported helper architecture.") };
        var expected = runtime.Parent?.Parent?.Parent?.FullName;
        if (runtime.Parent?.Name != "Microsoft.NETCore.App" || runtime.Parent.Parent?.Name != "shared" ||
            expected is null || !Version.TryParse(runtime.Name, out _))
            throw new IOException("Cannot identify the runtime hosting this process.");
        if (expected.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(expected, Path.Combine(root, "runtime", arch), StringComparison.OrdinalIgnoreCase))
            throw new IOException("Private runtime layout does not match the running process.");
        if (command.UseShellExecute) throw new IOException("Private runtime helpers require direct process creation.");
        if (!Path.IsPathFullyQualified(command.FileName)) throw new IOException("An absolute helper path is required.");
        using var stream = File.OpenRead(command.FileName);
        using var reader = new PEReader(stream);
        var headers = reader.PEHeaders;
        var machine = architecture == Architecture.X86 ? Machine.I386 : Machine.Amd64;
        var magic = architecture == Architecture.X86 ? PEMagic.PE32 : PEMagic.PE32Plus;
        if (headers.CoffHeader.Machine != machine || headers.PEHeader?.Magic != magic ||
            (headers.CoffHeader.Characteristics & Characteristics.ExecutableImage) == 0 ||
            (headers.CoffHeader.Characteristics & Characteristics.Dll) != 0)
            throw new IOException("Helper executable must match the selected private runtime architecture.");
        command.Environment["DOTNET_ROOT"] = expected;
        command.Environment["DOTNET_ROOT_" + arch.ToUpperInvariant()] = expected;
        return expected;
    }
}
