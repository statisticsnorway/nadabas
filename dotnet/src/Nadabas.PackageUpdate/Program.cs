using Nadabas.PackageUpdate;
using System.Text.Json;

try
{
    if (!OperatingSystem.IsWindows()) throw new IOException("This updater requires Windows.");
    if (args.Length is 2 or 3 && args[0] == "check-download" && (args.Length == 2 || args[2] == "--manual"))
    {
        Console.WriteLine(JsonSerializer.Serialize(await AutomaticUpdates.Check(args[1], args.Length == 3)));
        return 0;
    }
    if (args.Length is 2 or 3 && args[0] == "verify" && (args.Length == 2 || args[2] == "--smoke"))
    {
        Console.WriteLine(JsonSerializer.Serialize(new { status = "verified", version = PackageJob.VerifyPackage(PackageRules.SafePath(args[1]), args.Length == 3) }));
        return 0;
    }
    if (args.Length is 4 or 5 && args[0] == "stage" && (args.Length == 4 || args[4] == "--smoke"))
    {
        Console.WriteLine(JsonSerializer.Serialize(PackageJob.Stage(args[1], args[2], args[3], args.Length == 5)));
        return 0;
    }
    if (args.Length is 3 or 4 && args[0] == "stage-for-excel" && (args.Length == 3 || args[3] == "--smoke"))
    {
        Console.WriteLine(JsonSerializer.Serialize(PlatformHandoff.Stage(args[1], args[2], args.Length == 4)));
        return 0;
    }
    if (args.Length == 4 && args[0] == "recover-for-excel")
    {
        Console.WriteLine(JsonSerializer.Serialize(await PlatformHandoff.Recover(args[1], args[2], args[3])));
        return 0;
    }
    if (args.Length == 3 && args[0] == "inspect-for-excel")
    {
        Console.WriteLine(JsonSerializer.Serialize(PlatformHandoff.Verify(args[1], args[2])));
        return 0;
    }
    if (args.Length == 3 && args[0] == "cancel-for-excel")
    {
        PlatformHandoff.Cancel(args[1], args[2]);
        Console.WriteLine("Cancelled");
        return 0;
    }
    if (args.Length == 5 && args[0] == "wait-apply" && int.TryParse(args[3], out var pid) && long.TryParse(args[4], out var ticks))
    {
        await PlatformHandoff.WaitAndApply(args[1], args[2], pid, ticks);
        return 0;
    }
    if (args.Length == 2 && args[0] == "apply")
    {
        Console.WriteLine(JsonSerializer.Serialize(await PackageJob.Apply(args[1]), new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull }));
        return 0;
    }
    Console.Error.WriteLine("INTERNAL TEST: stage <installed-root> <signed-installer> <new-external-cache> [--smoke] | stage-for-excel <installed-root> <signed-installer> [--smoke] | apply <external-cache>");
    return 2;
}
catch (Exception ex)
{
    Console.Error.WriteLine(JsonSerializer.Serialize(new { status = "failed", error = ex.Message }));
    return 1;
}
