using Nadabas.Setup;

namespace Nadabas.Setup.Tests;

public sealed class LegacyMigrationTests
{
    [Fact]
    public void OnlyExactStartupPathIsInterpretedAsACandidate()
    {
        var file = Path.Combine(Path.GetTempPath(), "My AddIns", "NADABAS.XLAM");
        Assert.Equal(file, LegacyMigration.RegistrationPath("Options", "OPEN1", $"/R \"{file}\""));
        Assert.Equal(file, LegacyMigration.RegistrationPath("Add-in Manager", file, null));
        Assert.Null(LegacyMigration.RegistrationPath("Options", "OPENOTHER", file));
        Assert.Null(LegacyMigration.RegistrationPath("Options", "OPEN", $"/R \"{file}\" /other"));
    }

    [Fact]
    public void ACustomAddinIsPreservedAndIdentifiedByContent()
    {
        var directory = Directory.CreateTempSubdirectory("nadabas-legacy-");
        var path = Path.Combine(directory.FullName, "NADABAS.XLAM");
        try
        {
            File.WriteAllText(path, "local unpublished edits");
            var observed = LegacyMigration.Inspect(path, ["user-addins-directory"], Path.Combine(directory.FullName, "managed"));
            Assert.Equal("unknown-or-custom-content", observed.Classification);
            Assert.Equal(64, observed.Sha256!.Length);
            Assert.Equal("local unpublished edits", File.ReadAllText(path));
        }
        finally { File.Delete(path); directory.Delete(); }
    }
}
