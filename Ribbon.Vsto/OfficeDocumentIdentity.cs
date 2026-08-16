using System;
using System.Security.Cryptography;
using System.Text;

namespace Ribbon.Vsto
{
    public static class OfficeDocumentIdentity
    {
        public static string Get(string hostKind, string documentNameOrPath)
        {
            return Get(hostKind, documentNameOrPath, "document");
        }

        public static string Get(string hostKind, string value, string label)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            var prefix = string.IsNullOrWhiteSpace(hostKind) ? "office" : hostKind.Trim().ToLowerInvariant();
            var input = prefix + "|" + value.Trim().ToUpperInvariant();
            using (var algorithm = SHA256.Create())
            {
                var hash = algorithm.ComputeHash(Encoding.UTF8.GetBytes(input));
                var builder = new StringBuilder(hash.Length * 2);
                foreach (var byteValue in hash) builder.Append(byteValue.ToString("x2"));
                return prefix + "-" + label + "-" + builder;
            }
        }
    }
}
