using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Command.Tests
{
    [TestClass]
    public sealed class TranslationServiceTests
    {
        [TestMethod]
        public async Task DeepLUsesFreeEndpointAndParsesSuccessfulResponse()
        {
            var handler = new RecordingHandler(request => Json(
                HttpStatusCode.OK,
                "{\"translations\":[{\"detected_source_language\":\"DE\",\"text\":\"Hello\"}]}"));
            var service = new DeepLService("secret:fx", new HttpClient(handler));

            string result = await service.TranslateAsync("Hallo", "EN-US");

            Assert.AreEqual("Hello", result);
            Assert.AreEqual("https://api-free.deepl.com/v2/translate", handler.RequestUri.AbsoluteUri);
            StringAssert.Contains(handler.Body, "auth_key=secret%3Afx");
            StringAssert.Contains(handler.Body, "text=Hallo");
            StringAssert.Contains(handler.Body, "target_lang=EN-US");
        }

        [TestMethod]
        public async Task DeepLRejectsMissingKeyAndFailedOrMalformedResponses()
        {
            var unused = new RecordingHandler(_ => throw new InvalidOperationException("must not send"));
            Assert.IsNull(await new DeepLService(null, new HttpClient(unused)).TranslateAsync("Hallo", "EN"));
            Assert.AreEqual(0, unused.CallCount);

            var failed = new RecordingHandler(_ => Json((HttpStatusCode)429, "rate limited"));
            Assert.IsNull(await new DeepLService("key", new HttpClient(failed)).TranslateAsync("Hallo", "EN"));

            var malformed = new RecordingHandler(_ => Json(HttpStatusCode.OK, "{bad"));
            Assert.IsNull(await new DeepLService("key", new HttpClient(malformed)).TranslateAsync("Hallo", "EN"));
        }

        [TestMethod]
        public async Task LibreTranslateNormalizesTargetAndIncludesOptionalKey()
        {
            var handler = new RecordingHandler(_ => Json(HttpStatusCode.OK, "{\"translatedText\":\"Bonjour\"}"));
            var service = new LibreTranslateService("https://translate.example/api", "token", new HttpClient(handler));

            string result = await service.TranslateAsync("Hello", "FR-FR");

            Assert.AreEqual("Bonjour", result);
            Assert.AreEqual("https://translate.example/api/translate", handler.RequestUri.AbsoluteUri);
            StringAssert.Contains(handler.Body, "\"target\":\"fr\"");
            StringAssert.Contains(handler.Body, "\"api_key\":\"token\"");
        }

        [TestMethod]
        public async Task LibreTranslateHandlesApiErrorsTransportErrorsAndMalformedJson()
        {
            var apiError = new RecordingHandler(_ => Json(HttpStatusCode.OK, "{\"error\":\"unsupported\"}"));
            Assert.IsNull(await new LibreTranslateService(null, null, new HttpClient(apiError)).TranslateAsync("x", "EN"));

            var malformed = new RecordingHandler(_ => Json(HttpStatusCode.OK, "not-json"));
            Assert.IsNull(await new LibreTranslateService(null, null, new HttpClient(malformed)).TranslateAsync("x", "EN"));

            var transportError = new RecordingHandler(_ => throw new HttpRequestException("offline"));
            Assert.IsNull(await new LibreTranslateService(null, null, new HttpClient(transportError)).TranslateAsync("x", "EN"));
        }

        [TestMethod]
        public void LibreTranslateRequiresHttpsExceptForLoopbackDevelopmentServers()
        {
            Assert.AreEqual("https://translate.example/api/", LibreTranslateService.NormalizeBaseUrl("https://translate.example/api"));
            Assert.AreEqual("http://localhost:5000/api/", LibreTranslateService.NormalizeBaseUrl("http://localhost:5000/api"));
            Assert.ThrowsException<ArgumentException>(() => LibreTranslateService.NormalizeBaseUrl("http://translate.example/api"));
            Assert.ThrowsException<ArgumentException>(() => LibreTranslateService.NormalizeBaseUrl("ftp://translate.example/api"));
            Assert.ThrowsException<ArgumentException>(() => LibreTranslateService.NormalizeBaseUrl("https://user:password@translate.example/api"));
            Assert.ThrowsException<ArgumentException>(() => LibreTranslateService.NormalizeBaseUrl("https://translate.example/api?tenant=secret"));
        }

        [TestMethod]
        public async Task OpenAITranslationBuildsAuthenticatedRequestAndTrimsResult()
        {
            var handler = new RecordingHandler(_ => Json(
                HttpStatusCode.OK,
                "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"content\":\"  Hello  \"}}]}"));
            var service = new LLMTranslationService(new HttpClient(handler));

            string result = await service.TranslateWithOpenAIAsync("Hallo", "English", "secret", "model-x");

            Assert.AreEqual("Hello", result);
            Assert.AreEqual("https://api.openai.com/v1/chat/completions", handler.RequestUri.AbsoluteUri);
            Assert.AreEqual("Bearer", handler.AuthorizationScheme);
            Assert.AreEqual("secret", handler.AuthorizationParameter);
            StringAssert.Contains(handler.Body, "\"model\":\"model-x\"");
            StringAssert.Contains(handler.Body, "\"content\":\"Hallo\"");
        }

        [TestMethod]
        public async Task ClaudeTranslationBuildsVersionedRequestAndHandlesFailures()
        {
            var handler = new RecordingHandler(_ => Json(
                HttpStatusCode.OK,
                "{\"content\":[{\"text\":\"  Salut  \"}]}"));
            var service = new LLMTranslationService(new HttpClient(handler));

            string result = await service.TranslateWithClaudeAsync("Hello", "French", "secret", "claude-x");

            Assert.AreEqual("Salut", result);
            Assert.AreEqual("secret", handler.ApiKey);
            Assert.AreEqual("2023-06-01", handler.AnthropicVersion);
            StringAssert.Contains(handler.Body, "\"model\":\"claude-x\"");

            var failed = new RecordingHandler(_ => Json(HttpStatusCode.BadRequest, "bad"));
            Assert.IsNull(await new LLMTranslationService(new HttpClient(failed))
                .TranslateWithClaudeAsync("x", "French", "key", "model"));
            Assert.IsNull(await service.TranslateWithClaudeAsync("x", "French", null, "model"));
        }

        [TestMethod]
        public async Task AllTranslationProvidersRejectOversizedResponseBodies()
        {
            Assert.IsNull(await new DeepLService("key", ClientWithUnknownLengthBody(OversizedBody()))
                .TranslateAsync("x", "EN"));
            Assert.IsNull(await new LibreTranslateService(null, null, ClientWithUnknownLengthBody(OversizedBody()))
                .TranslateAsync("x", "EN"));
            Assert.IsNull(await new LLMTranslationService(ClientWithUnknownLengthBody(OversizedBody()))
                .TranslateWithOpenAIAsync("x", "English", "key", "model"));
            Assert.IsNull(await new LLMTranslationService(ClientWithUnknownLengthBody(OversizedBody()))
                .TranslateWithClaudeAsync("x", "English", "key", "model"));
        }

        [TestMethod]
        public async Task AllTranslationProvidersRejectOversizedParsedTranslations()
        {
            string output = new string('x', BoundedHttpResponseReader.MaxTranslationCharacters + 1);

            Assert.IsNull(await new DeepLService("key", ClientWithJson(
                    "{\"translations\":[{\"text\":\"" + output + "\"}]}"))
                .TranslateAsync("x", "EN"));
            Assert.IsNull(await new LibreTranslateService(null, null, ClientWithJson(
                    "{\"translatedText\":\"" + output + "\"}"))
                .TranslateAsync("x", "EN"));
            Assert.IsNull(await new LLMTranslationService(ClientWithJson(
                    "{\"choices\":[{\"message\":{\"content\":\"" + output + "\"}}]}"))
                .TranslateWithOpenAIAsync("x", "English", "key", "model"));
            Assert.IsNull(await new LLMTranslationService(ClientWithJson(
                    "{\"content\":[{\"text\":\"" + output + "\"}]}"))
                .TranslateWithClaudeAsync("x", "English", "key", "model"));
        }

        [TestMethod]
        public async Task AllTranslationProvidersRejectTruncatedResponses()
        {
            Assert.IsNull(await new DeepLService("key", ClientWithTruncatedBody("{}"))
                .TranslateAsync("x", "EN"));
            Assert.IsNull(await new LibreTranslateService(null, null, ClientWithTruncatedBody("{}"))
                .TranslateAsync("x", "EN"));
            Assert.IsNull(await new LLMTranslationService(ClientWithTruncatedBody("{}"))
                .TranslateWithOpenAIAsync("x", "English", "key", "model"));
            Assert.IsNull(await new LLMTranslationService(ClientWithTruncatedBody("{}"))
                .TranslateWithClaudeAsync("x", "English", "key", "model"));
        }

        [TestMethod]
        public async Task AllTranslationProvidersApplyAnExplicitRequestTimeout()
        {
            TimeSpan timeout = TimeSpan.FromMilliseconds(25);

            Assert.IsNull(await new DeepLService("key", ClientThatWaitsForCancellation(), timeout)
                .TranslateAsync("x", "EN"));
            Assert.IsNull(await new LibreTranslateService(null, null, ClientThatWaitsForCancellation(), timeout)
                .TranslateAsync("x", "EN"));
            Assert.IsNull(await new LLMTranslationService(ClientThatWaitsForCancellation(), timeout)
                .TranslateWithOpenAIAsync("x", "English", "key", "model"));
            Assert.IsNull(await new LLMTranslationService(ClientThatWaitsForCancellation(), timeout)
                .TranslateWithClaudeAsync("x", "English", "key", "model"));
        }

        [TestMethod]
        public async Task OpenAIAndClaudeRejectMalformedJson()
        {
            Assert.IsNull(await new LLMTranslationService(ClientWithJson("{bad"))
                .TranslateWithOpenAIAsync("x", "English", "key", "model"));
            Assert.IsNull(await new LLMTranslationService(ClientWithJson("{bad"))
                .TranslateWithClaudeAsync("x", "English", "key", "model"));
        }

        private static HttpResponseMessage Json(HttpStatusCode status, string body)
        {
            return new HttpResponseMessage(status)
            {
                Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json")
            };
        }

        private static HttpClient ClientWithJson(string body)
        {
            return new HttpClient(new RecordingHandler(_ => Json(HttpStatusCode.OK, body)));
        }

        private static HttpClient ClientWithUnknownLengthBody(string body)
        {
            return new HttpClient(new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(new MemoryStream(Encoding.UTF8.GetBytes(body)))
            }));
        }

        private static HttpClient ClientWithTruncatedBody(string body)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(body);
            var content = new StreamContent(new MemoryStream(bytes));
            content.Headers.ContentLength = bytes.Length + 1;
            return new HttpClient(new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = content
            }));
        }

        private static HttpClient ClientThatWaitsForCancellation()
        {
            return new HttpClient(new AsyncHandler(async cancellationToken =>
            {
                await Task.Delay(Timeout.Infinite, cancellationToken);
                return Json(HttpStatusCode.OK, "{}");
            }));
        }

        private static string OversizedBody()
        {
            return new string('x', BoundedHttpResponseReader.MaxResponseBytes + 1);
        }

        private sealed class RecordingHandler : HttpMessageHandler
        {
            private readonly Func<HttpRequestMessage, HttpResponseMessage> _response;

            public RecordingHandler(Func<HttpRequestMessage, HttpResponseMessage> response)
            {
                _response = response;
            }

            public int CallCount { get; private set; }
            public Uri RequestUri { get; private set; }
            public string Body { get; private set; }
            public string AuthorizationScheme { get; private set; }
            public string AuthorizationParameter { get; private set; }
            public string ApiKey { get; private set; }
            public string AnthropicVersion { get; private set; }

            protected override async Task<HttpResponseMessage> SendAsync(
                HttpRequestMessage request,
                CancellationToken cancellationToken)
            {
                CallCount++;
                RequestUri = request.RequestUri;
                Body = request.Content == null ? null : await request.Content.ReadAsStringAsync();
                AuthorizationScheme = request.Headers.Authorization?.Scheme;
                AuthorizationParameter = request.Headers.Authorization?.Parameter;
                ApiKey = Header(request, "x-api-key");
                AnthropicVersion = Header(request, "anthropic-version");
                return _response(request);
            }

            private static string Header(HttpRequestMessage request, string name)
            {
                return request.Headers.TryGetValues(name, out var values) ? System.Linq.Enumerable.Single(values) : null;
            }
        }

        private sealed class AsyncHandler : HttpMessageHandler
        {
            private readonly Func<CancellationToken, Task<HttpResponseMessage>> _response;

            internal AsyncHandler(Func<CancellationToken, Task<HttpResponseMessage>> response)
            {
                _response = response;
            }

            protected override Task<HttpResponseMessage> SendAsync(
                HttpRequestMessage request,
                CancellationToken cancellationToken)
            {
                return _response(cancellationToken);
            }
        }
    }
}
