using System;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free placeholder generation for slide-text anonymization (QA-S9 dialogs/workflows
    /// wave): each letter or digit position of the original text is replaced with a random
    /// lorem-vocabulary letter while whitespace, punctuation, symbol positions, string length
    /// and per-position letter case are preserved. Production implementation used by
    /// AnonymizeFeature; pass a seeded Random for reproducible output.
    /// </summary>
    public static class AnonymizePlaceholderPolicy
    {
        public const string PlaceholderCharacters = "loremipsumdolorsitametconsectetur";

        public static string GeneratePlaceholder(string original, Random random)
        {
            char[] result = original.ToCharArray();
            for (int index = 0; index < result.Length; index++)
            {
                if (!char.IsLetterOrDigit(original[index])) continue;
                char replacement = PlaceholderCharacters[random.Next(PlaceholderCharacters.Length)];
                result[index] = char.IsUpper(original[index]) ? char.ToUpperInvariant(replacement) : replacement;
            }
            return new string(result);
        }
    }
}
