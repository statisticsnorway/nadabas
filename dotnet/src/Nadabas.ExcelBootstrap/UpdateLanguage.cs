using System.Text.Json;
using Microsoft.Win32;

namespace Nadabas.ExcelBootstrap;

internal static class UpdateLanguage
{
    private static readonly Dictionary<string, Dictionary<string, string>> Tables = Load();

    private static Dictionary<string, Dictionary<string, string>> Load()
    {
        using var stream = typeof(UpdateLanguage).Assembly.GetManifestResourceStream("Nadabas.UpdateText.json")!;
        return JsonSerializer.Deserialize<Dictionary<string, Dictionary<string, string>>>(stream)!;
    }

    internal static string CurrentCode()
    {
        try
        {
            var code = Registry.GetValue(@"HKEY_CURRENT_USER\Software\VB and VBA Program Settings\NADABAS\Language", "Code", "3")?.ToString();
            return code is "4" or "5" or "6" ? code : "3";
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException) { return "3"; }
    }

    internal static string Text(string code, string key) => Tables[code][key];
}
