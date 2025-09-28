using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace InventorApp.API.Services
{
    public class EncryptionService
    {
        private readonly byte[] _key;
        private readonly byte[] _iv;

        public EncryptionService(IConfiguration configuration)
        {
            var keyBase64 = configuration["Encryption:KeyBase64"] ?? string.Empty;
            var ivBase64 = configuration["Encryption:IVBase64"] ?? string.Empty;

            if (string.IsNullOrWhiteSpace(keyBase64) || string.IsNullOrWhiteSpace(ivBase64))
            {
                throw new InvalidOperationException("Encryption key/IV not configured. Set Encryption:KeyBase64 and Encryption:IVBase64.");
            }

            _key = Convert.FromBase64String(keyBase64);
            _iv = Convert.FromBase64String(ivBase64);

            if (_key.Length != 32)
            {
                throw new InvalidOperationException("Encryption key must be 32 bytes (AES-256).");
            }
            if (_iv.Length != 16)
            {
                throw new InvalidOperationException("Encryption IV must be 16 bytes.");
            }
        }

        public string EncryptToBase64(string plainText)
        {
            if (string.IsNullOrEmpty(plainText)) return string.Empty;

            using var aesAlg = Aes.Create();
            aesAlg.Key = _key;
            aesAlg.IV = _iv;
            aesAlg.Mode = CipherMode.CBC;
            aesAlg.Padding = PaddingMode.PKCS7;

            using var encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
            using var msEncrypt = new MemoryStream();
            using (var csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
            using (var swEncrypt = new StreamWriter(csEncrypt, Encoding.UTF8))
            {
                swEncrypt.Write(plainText);
            }
            return Convert.ToBase64String(msEncrypt.ToArray());
        }

        public string DecryptFromBase64(string cipherTextBase64)
        {
            if (string.IsNullOrEmpty(cipherTextBase64)) return string.Empty;

            var cipherBytes = Convert.FromBase64String(cipherTextBase64);

            using var aesAlg = Aes.Create();
            aesAlg.Key = _key;
            aesAlg.IV = _iv;
            aesAlg.Mode = CipherMode.CBC;
            aesAlg.Padding = PaddingMode.PKCS7;

            using var decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
            using var msDecrypt = new MemoryStream(cipherBytes);
            using var csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read);
            using var srDecrypt = new StreamReader(csDecrypt, Encoding.UTF8);
            return srDecrypt.ReadToEnd();
        }
    }
}

