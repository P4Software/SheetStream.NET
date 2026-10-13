using System;
using System.Security.Cryptography;
using System.Text;

namespace SheetStream.Internal
{
    /// <summary>
    /// Computes the ISO/IEC 29500 password verifier used by sheet and workbook protection: SHA-512 over a
    /// random salt plus the UTF-16LE password, then iterated <see cref="SpinCount"/> times with the iteration
    /// counter appended. This is what Excel 2013+ writes. The legacy 16-bit hash is trivially reversible and
    /// is deliberately not offered.
    /// </summary>
    internal static class PasswordHasher
    {
        /// <summary>Iteration count Excel uses by default.</summary>
        public const int SpinCount = 100000;

        /// <summary>Returns (hash, salt) as base64 strings.</summary>
        public static (string Hash, string Salt) Hash(string password)
        {
            if (password == null) throw new ArgumentNullException(nameof(password));
            var salt = new byte[16];
            using (var rng = RandomNumberGenerator.Create()) rng.GetBytes(salt);
            return (Convert.ToBase64String(Compute(password, salt, SpinCount)), Convert.ToBase64String(salt));
        }

        /// <summary>Computes the verifier for a known salt. Exposed internally for tests.</summary>
        public static byte[] Compute(string password, byte[] salt, int spinCount)
        {
            using (var sha = SHA512.Create())
            {
                var pwd = Encoding.Unicode.GetBytes(password);
                var buffer = new byte[salt.Length + pwd.Length];
                Buffer.BlockCopy(salt, 0, buffer, 0, salt.Length);
                Buffer.BlockCopy(pwd, 0, buffer, salt.Length, pwd.Length);
                var hash = sha.ComputeHash(buffer);

                // Each round hashes the previous hash followed by the little-endian 32-bit iteration number.
                var round = new byte[hash.Length + 4];
                for (var i = 0; i < spinCount; i++)
                {
                    Buffer.BlockCopy(hash, 0, round, 0, hash.Length);
                    round[hash.Length] = (byte)i;
                    round[hash.Length + 1] = (byte)(i >> 8);
                    round[hash.Length + 2] = (byte)(i >> 16);
                    round[hash.Length + 3] = (byte)(i >> 24);
                    hash = sha.ComputeHash(round);
                }
                return hash;
            }
        }
    }
}
