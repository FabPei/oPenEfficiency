using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 dialogs/workflows-wave coverage for slide-text anonymization
    /// (QA9-BtnAnonymize-002): the lorem placeholder generation with its structure,
    /// vocabulary and case-preservation contract, used by AnonymizeFeature.
    /// </summary>
    [TestClass]
    public sealed class AnonymizePlaceholderPolicyTests
    {
        [TestMethod]
        public void PreservesLengthAndNonLetterPositions()
        {
            string placeholder = AnonymizePlaceholderPolicy.GeneratePlaceholder("Hello, Welt 42!", new Random(7));
            Assert.AreEqual(15, placeholder.Length);
            Assert.AreEqual(',', placeholder[5]);
            Assert.AreEqual(' ', placeholder[6]);
            Assert.AreEqual(' ', placeholder[11]);
            Assert.AreEqual('!', placeholder[14]);
        }

        [TestMethod]
        public void ReplacesLettersAndDigitsWithLoremVocabularyLetters()
        {
            string placeholder = AnonymizePlaceholderPolicy.GeneratePlaceholder("42", new Random(11));
            Assert.AreEqual(2, placeholder.Length);
            foreach (char c in placeholder)
            {
                Assert.IsTrue(char.IsLetter(c), $"'{c}' is not a letter; digits are replaced with letters.");
                Assert.IsTrue(AnonymizePlaceholderPolicy.PlaceholderCharacters.IndexOf(char.ToLowerInvariant(c)) >= 0,
                    $"'{c}' is outside the lorem vocabulary.");
            }
        }

        [TestMethod]
        public void PreservesLetterCasePerPosition()
        {
            string upper = AnonymizePlaceholderPolicy.GeneratePlaceholder("ABC", new Random(3));
            Assert.AreEqual(3, upper.Length);
            CollectionAssert.AreEqual(
                new[] { true, true, true },
                new[] { char.IsUpper(upper[0]), char.IsUpper(upper[1]), char.IsUpper(upper[2]) });

            string lower = AnonymizePlaceholderPolicy.GeneratePlaceholder("abc", new Random(3));
            CollectionAssert.AreEqual(
                new[] { false, false, false },
                new[] { char.IsUpper(lower[0]), char.IsUpper(lower[1]), char.IsUpper(lower[2]) });
        }

        [TestMethod]
        public void IsDeterministicForASeedAndVariesAcrossSeeds()
        {
            string first = AnonymizePlaceholderPolicy.GeneratePlaceholder("Confidential quarterly report", new Random(42));
            string second = AnonymizePlaceholderPolicy.GeneratePlaceholder("Confidential quarterly report", new Random(42));
            Assert.AreEqual(first, second);

            string third = AnonymizePlaceholderPolicy.GeneratePlaceholder("Confidential quarterly report", new Random(43));
            Assert.AreNotEqual(first, third);
        }

        [TestMethod]
        public void HandlesEmptyAndSymbolOnlyStrings()
        {
            Assert.AreEqual(string.Empty, AnonymizePlaceholderPolicy.GeneratePlaceholder(string.Empty, new Random(1)));
            Assert.AreEqual("...!?-", AnonymizePlaceholderPolicy.GeneratePlaceholder("...!?-", new Random(1)));
        }
    }
}
