using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace oPenEfficiency.Utils
{
    /// <summary>
    /// Removes configured credentials and common credential representations from
    /// local diagnostics before they reach a log file or debug output.
    /// </summary>
    public static class SecretRedactor
    {
        private const string Replacement = "[REDACTED]";
        private static readonly object SyncRoot = new object();
        private static string[] _knownSecrets = Array.Empty<string>();

        private static readonly Regex JsonCredentialPattern = new Regex(
            "(\\\"(?:DeepLApiKey|OpenAIApiKey|ClaudeApiKey|LibreTranslateApiKey|api[_-]?key|access[_-]?token|secret)\\\"\\s*:\\s*\\\")((?:\\\\.|[^\\\"\\\\])*)(\\\")",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex HeaderCredentialPattern = new Regex(
            "((?:Authorization\\s*:\\s*Bearer|x-api-key|api[_-]?key|auth_key)\\s*[=:]?\\s*)([^\\s,;&]+)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        public static void SetKnownSecrets(IEnumerable<string> secrets)
        {
            string[] replacement = (secrets ?? Enumerable.Empty<string>())
                .Where(secret => !string.IsNullOrEmpty(secret))
                .Distinct(StringComparer.Ordinal)
                .OrderByDescending(secret => secret.Length)
                .ToArray();

            lock (SyncRoot)
            {
                _knownSecrets = replacement;
            }
        }

        public static void RegisterSecrets(IEnumerable<string> secrets)
        {
            string[] additions = (secrets ?? Enumerable.Empty<string>())
                .Where(secret => !string.IsNullOrEmpty(secret))
                .ToArray();
            if (additions.Length == 0) return;

            lock (SyncRoot)
            {
                _knownSecrets = _knownSecrets
                    .Concat(additions)
                    .Distinct(StringComparer.Ordinal)
                    .OrderByDescending(secret => secret.Length)
                    .ToArray();
            }
        }

        public static string Redact(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;

            string redacted = JsonCredentialPattern.Replace(value, "$1" + Replacement + "$3");
            redacted = HeaderCredentialPattern.Replace(redacted, "$1" + Replacement);

            string[] secrets;
            lock (SyncRoot)
            {
                secrets = _knownSecrets;
            }

            foreach (string secret in secrets)
            {
                redacted = redacted.Replace(secret, Replacement);
            }

            return redacted;
        }
    }
}
