using Chaos.NaCl;
using Nadabas.FeedSigning;
using Nadabas.PackageUpdate;
using System.Security.Cryptography;
using System.Text.Json.Nodes;

namespace Nadabas.Setup.Tests;

public sealed class ReleaseKeyTests
{
    [Fact]
    public void BridgeAcceptsBothKeysAndRevocationOverridesPrimaryKey()
    {
        Ed25519.KeyPairFromSeed(out var oldPublic, out var oldPrivate, RandomNumberGenerator.GetBytes(32));
        Ed25519.KeyPairFromSeed(out var newPublic, out var newPrivate, RandomNumberGenerator.GetBytes(32));
        byte[] payload = [5, 6, 7];
        var oldKey = Convert.ToBase64String(oldPublic);
        var newKey = Convert.ToBase64String(newPublic);
        var oldSignature = Convert.ToBase64String(Ed25519.Sign(payload, oldPrivate));
        var newSignature = Convert.ToBase64String(Ed25519.Sign(payload, newPrivate));
        var bridge = new UpdateChannel("unused", oldKey, [newKey]);
        Assert.True(bridge.Verify(oldSignature, payload));
        Assert.True(bridge.Verify(newSignature, payload));
        var retired = bridge with { RevokedPublicKeys = [oldKey] };
        Assert.False(retired.Verify(oldSignature, payload));
        Assert.True(retired.Verify(newSignature, payload));
        Assert.Throws<IOException>(() => (bridge with { RevokedPublicKeys = [oldKey, newKey] }).TrustedKeys());
        Assert.Throws<IOException>(() => (bridge with { AdditionalPublicKeys = [oldKey + "\n"] }).TrustedKeys());
    }

    [Fact]
    public void PortableBackupRecoversSameSigningIdentityAndRejectsWrongPasswordOrTampering()
    {
        var seed = RandomNumberGenerator.GetBytes(32);
        const string password = "ephemeral test password only";
        var backup = PortableSeedBackup.Encrypt(seed, password);
        var recovered = PortableSeedBackup.Decrypt(backup, password);
        Assert.Equal(seed, recovered);
        Ed25519.KeyPairFromSeed(out var originalPublic, out _, seed);
        Ed25519.KeyPairFromSeed(out var restoredPublic, out var restoredPrivate, recovered);
        byte[] challenge = [2, 5, 8];
        Assert.Equal(originalPublic, restoredPublic);
        Assert.True(Ed25519.Verify(Ed25519.Sign(challenge, restoredPrivate), challenge, originalPublic));
        Assert.ThrowsAny<CryptographicException>(() => PortableSeedBackup.Decrypt(backup, "wrong password"));
        var document = JsonNode.Parse(backup)!;
        var ciphertext = Convert.FromBase64String(document["Ciphertext"]!.GetValue<string>());
        ciphertext[0] ^= 1;
        document["Ciphertext"] = Convert.ToBase64String(ciphertext);
        Assert.ThrowsAny<CryptographicException>(() =>
            PortableSeedBackup.Decrypt(System.Text.Encoding.UTF8.GetBytes(document.ToJsonString()), password));
        Assert.Throws<ArgumentException>(() => PortableSeedBackup.Encrypt(seed, "short"));
    }
}
