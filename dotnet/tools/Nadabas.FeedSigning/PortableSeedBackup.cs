using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace Nadabas.FeedSigning;

// Portable recovery envelope. DPAPI continues to protect the working copy.
internal static class PortableSeedBackup
{
    private const int Iterations = 600_000;
    private static readonly byte[] Context = Encoding.UTF8.GetBytes("NADABAS feed seed backup v1");
    private sealed record Envelope(int SchemaVersion, byte[] Salt, byte[] Nonce, byte[] Ciphertext, byte[] Tag);

    internal static byte[] Encrypt(byte[] seed, ReadOnlySpan<char> password)
    {
        if (seed.Length != 32) throw new CryptographicException("Invalid signing seed.");
        if (password.Length < 16) throw new ArgumentException("Use a recovery passphrase of at least 16 characters.");
        var salt = RandomNumberGenerator.GetBytes(16);
        var nonce = RandomNumberGenerator.GetBytes(12);
        var key = Rfc2898DeriveBytes.Pbkdf2(password, salt, Iterations, HashAlgorithmName.SHA256, 32);
        try
        {
            var ciphertext = new byte[32];
            var tag = new byte[16];
            using var aes = new AesGcm(key, 16);
            aes.Encrypt(nonce, seed, ciphertext, tag, Context);
            return JsonSerializer.SerializeToUtf8Bytes(new Envelope(1, salt, nonce, ciphertext, tag));
        }
        finally { CryptographicOperations.ZeroMemory(key); }
    }

    internal static byte[] Decrypt(byte[] backup, ReadOnlySpan<char> password)
    {
        if (backup.Length > 4096) throw new CryptographicException("Invalid backup size.");
        var data = JsonSerializer.Deserialize<Envelope>(backup) ?? throw new CryptographicException("Empty backup.");
        if (data.SchemaVersion != 1 || data.Salt?.Length != 16 || data.Nonce?.Length != 12 ||
            data.Ciphertext?.Length != 32 || data.Tag?.Length != 16)
            throw new CryptographicException("Invalid backup format.");
        var key = Rfc2898DeriveBytes.Pbkdf2(password, data.Salt, Iterations, HashAlgorithmName.SHA256, 32);
        var seed = new byte[32];
        try
        {
            using var aes = new AesGcm(key, 16);
            aes.Decrypt(data.Nonce, data.Ciphertext, data.Tag, seed, Context);
            return seed;
        }
        catch
        {
            CryptographicOperations.ZeroMemory(seed);
            throw;
        }
        finally { CryptographicOperations.ZeroMemory(key); }
    }
}
