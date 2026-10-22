using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class SecretRedactorTests
    {
        [TestCleanup]
        public void TearDown()
        {
            SecretRedactor.SetKnownSecrets(null);
        }

        [TestMethod]
        public void RedactsKnownSecretsAndCredentialRepresentations()
        {
            SecretRedactor.SetKnownSecrets(new[] { "exact-super-secret" });

            string redacted = SecretRedactor.Redact(
                "exact-super-secret Authorization: Bearer bearer-value " +
                "x-api-key=header-value auth_key=query-value " +
                "{\"OpenAIApiKey\":\"json-value\"}");

            Assert.IsFalse(redacted.Contains("exact-super-secret"));
            Assert.IsFalse(redacted.Contains("bearer-value"));
            Assert.IsFalse(redacted.Contains("header-value"));
            Assert.IsFalse(redacted.Contains("query-value"));
            Assert.IsFalse(redacted.Contains("json-value"));
            StringAssert.Contains(redacted, "[REDACTED]");
        }

        [TestMethod]
        public void RedactsEscapedJsonValuesAndRetainsPreviouslyRegisteredSecrets()
        {
            SecretRedactor.SetKnownSecrets(new[] { "first-secret" });
            SecretRedactor.RegisterSecrets(new[] { "second-secret" });

            string redacted = SecretRedactor.Redact(
                "first-secret second-secret {\"api_key\":\"part\\\"still-secret\"}");

            Assert.IsFalse(redacted.Contains("first-secret"));
            Assert.IsFalse(redacted.Contains("second-secret"));
            Assert.IsFalse(redacted.Contains("still-secret"));
        }

        [TestMethod]
        public void OrdinaryDiagnosticTextRemainsReadable()
        {
            const string message = "Translation request failed with status 429.";
            Assert.AreEqual(message, SecretRedactor.Redact(message));
        }
    }
}
