using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace oPenEfficiency.Services
{
    [DataContract]
    public class LibreTranslateResponse
    {
        [DataMember(Name = "translatedText")]
        public string TranslatedText { get; set; }
        
        [DataMember(Name = "error")]
        public string Error { get; set; }
    }

    [DataContract]
    public class LibreTranslateRequest
    {
        [DataMember(Name = "q")]
        public string Text { get; set; }
        [DataMember(Name = "source")]
        public string Source { get; set; }
        [DataMember(Name = "target")]
        public string Target { get; set; }
        [DataMember(Name = "format")]
        public string Format { get; set; }
        [DataMember(Name = "api_key", EmitDefaultValue = false)]
        public string ApiKey { get; set; }
    }

    public class LibreTranslateService
    {
        private readonly string _baseUrl;
        private readonly string _apiKey;
        private readonly HttpClient _httpClient;
        private readonly TimeSpan _requestTimeout;
        private static readonly HttpClient SharedHttpClient = new HttpClient();

        public LibreTranslateService(string baseUrl, string apiKey = null)
            : this(baseUrl, apiKey, SharedHttpClient)
        {
        }

        internal LibreTranslateService(string baseUrl, string apiKey, HttpClient httpClient)
            : this(baseUrl, apiKey, httpClient, BoundedHttpResponseReader.DefaultTimeout)
        {
        }

        internal LibreTranslateService(string baseUrl, string apiKey, HttpClient httpClient, TimeSpan requestTimeout)
        {
            _baseUrl = NormalizeBaseUrl(baseUrl);
            _apiKey = apiKey;
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _requestTimeout = requestTimeout;
        }

        internal static string NormalizeBaseUrl(string baseUrl)
        {
            string candidate = string.IsNullOrWhiteSpace(baseUrl)
                ? "https://libretranslate.com/"
                : baseUrl.Trim();
            if (!Uri.TryCreate(candidate, UriKind.Absolute, out var uri))
                throw new ArgumentException("LibreTranslate Base URL must be an absolute HTTPS URL.", nameof(baseUrl));

            bool isHttps = string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase);
            bool isLoopbackHttp = string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) && uri.IsLoopback;
            if (!isHttps && !isLoopbackHttp)
                throw new ArgumentException("LibreTranslate Base URL must use HTTPS. HTTP is allowed only for a loopback server.", nameof(baseUrl));
            if (!string.IsNullOrEmpty(uri.UserInfo) || !string.IsNullOrEmpty(uri.Query) || !string.IsNullOrEmpty(uri.Fragment))
                throw new ArgumentException("LibreTranslate Base URL must not contain credentials, a query, or a fragment.", nameof(baseUrl));

            string normalized = uri.AbsoluteUri;
            return normalized.EndsWith("/", StringComparison.Ordinal) ? normalized : normalized + "/";
        }

        public async Task<string> TranslateAsync(string text, string targetLanguage)
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;

                string endpoint = _baseUrl + "translate";
                
                string target = targetLanguage;
                if (target.Contains("-")) target = target.Split('-')[0];
                target = target.ToLower();

                var requestData = new LibreTranslateRequest
                {
                    Text = text,
                    Source = "auto",
                    Target = target,
                    Format = "text",
                    ApiKey = string.IsNullOrEmpty(_apiKey) ? null : _apiKey
                };

                string jsonContent = JsonService.Serialize(requestData);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                BoundedHttpResponse response;
                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
                {
                    request.Content = content;
                    response = await BoundedHttpResponseReader.SendAsync(_httpClient, request, _requestTimeout);
                }

                if (!response.IsSuccessStatusCode)
                {
                    oPenEfficiency.Utils.ExceptionLogger.Log(
                        $"LibreTranslate returned HTTP {(int)response.StatusCode} ({response.StatusCode}). The response body was omitted because it may contain sensitive data.",
                        "LibreTranslateService.TranslateAsync");
                    return null;
                }

                var result = JsonService.Deserialize<LibreTranslateResponse>(response.Body);
                if (!string.IsNullOrEmpty(result?.Error))
                {
                     oPenEfficiency.Utils.ExceptionLogger.Log(
                         "LibreTranslate returned an application error. The response text was omitted because it may contain sensitive data.",
                         "LibreTranslateService.TranslateAsync");
                     return null;
                }

                return BoundedHttpResponseReader.ValidateTranslation(result?.TranslatedText);
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "LibreTranslateService.TranslateAsync");
            }

            return null;
        }
    }
}
