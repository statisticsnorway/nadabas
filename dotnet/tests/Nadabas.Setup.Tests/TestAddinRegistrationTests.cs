using Nadabas.Setup;

namespace Nadabas.Setup.Tests;

public sealed class TestAddinRegistrationTests
{
    private const string TestPath = @"C:\test\NADABAS.TEST.abc.xlam";

    [Fact]
    public void OnlyExactTestReferencesAreSelectedForCleanup()
    {
        Assert.True(TestAddinRegistration.RefersToTestFile(TestPath, "", TestPath));
        Assert.True(TestAddinRegistration.RefersToTestFile("OPEN2", "/R \"" + TestPath + "\"", TestPath));
        Assert.True(TestAddinRegistration.RefersToTestFile("OPEN", "\"" + TestPath + "\"", TestPath));
        Assert.False(TestAddinRegistration.RefersToTestFile("OPEN", @"C:\Users\user\AppData\Roaming\Microsoft\AddIns\NADABAS.XLAM", TestPath));
        Assert.False(TestAddinRegistration.RefersToTestFile(TestPath + ".backup", "", TestPath));
        Assert.False(TestAddinRegistration.RefersToTestFile("OTHER", TestPath, TestPath));
        Assert.False(TestAddinRegistration.RefersToTestFile("OPEN", TestPath + " extra", TestPath));
    }
}
