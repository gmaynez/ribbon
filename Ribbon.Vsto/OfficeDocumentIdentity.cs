using System;
using System.Security.Cryptography;
using System.Text;

namespace Ribbon.Vsto
{
    public static class OfficeDocumentIdentity
    {
        public static string Get(string hostKind, string documentNameOrPath)
        {
            if (string.IsNullOrWhiteSpace(documentNameOrPath)) return string.Empty;
            var prefix = string.IsNullOrWhiteSpace(hostKind) ? "office" : hostKind.Trim().ToLowerInvariant();
            var input = prefix + "|" + documentNameOrPath.Trim().ToUpperInvariant();
            using (var algorithm = SHA256.Create())
            {
                var hash = algorithm.ComputeHash(Encoding.UTF8.GetBytes(input));
                var builder = new StringBuilder(hash.Length * 2);
                foreach (var value in hash) builder.Append(value.ToString("x2"));
                return prefix + "-document-" + builder;
            }
        }
    }
}
