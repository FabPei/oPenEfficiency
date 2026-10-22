using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Threading.Tasks;

namespace oPenEfficiency.Services
{
    [DataContract]
    public class DeepLResponse
    {
        [DataMember(Name = "translations")]
        public List<DeepLTranslation> Translations { get; set; }
    }

    [DataContract]
    public class DeepLTranslation
    {
        [DataMember(Name = "detected_source_language")]
        public string DetectedSourceLanguage { get; set; }
        [DataMember(Name = "text")]
        public string Text { get; set; }
    }

    public class DeepLService
    {
        private readonly string _apiKey;
        private readonly HttpClient _httpClient;
        private readonly TimeSpan _requestTimeout;
        private static readonly HttpClient SharedHttpClient = new HttpClient();

        public DeepLService(string apiKey)
            : this(apiKey, SharedHttpClient)
        {
        }

        internal DeepLService(string apiKey, HttpClient httpClient)
            : this(apiKey, httpClient, BoundedHttpResponseReader.DefaultTimeout)
        {
        }

        internal DeepLService(string apiKey, HttpClient httpClient, TimeSpan requestTimeout)
        {
            _apiKey = apiKey;
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _requestTimeout = requestTimeout;
        }

        public async Task<string> TranslateAsync(string text, string targetLanguage)
        {
            if (string.IsNullOrEmpty(_apiKey)) return null;

            try
            {
                // Force TLS 1.2 for modern API compatibility in .NET 4.8
                System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;

                // DeepL Free API uses api-free.deepl.com
                // DeepL Pro API uses api.deepl.com
                // We check the key suffix to determine the endpoint (Free keys end with :fx)
                string endpoint = _apiKey.EndsWith(":fx") 
                    ? "https://api-free.deepl.com/v2/translate" 
                    : "https://api.deepl.com/v2/translate";

                var content = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("auth_key", _apiKey),
                    new KeyValuePair<string, string>("text", text),
                    new KeyValuePair<string, string>("target_lang", targetLanguage)
                });

                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
                {
                    request.Content = content;
                    var response = await BoundedHttpResponseReader.SendAsync(_httpClient, request, _requestTimeout);
                    if (!response.IsSuccessStatusCode)
                    {
                        oPenEfficiency.Utils.ExceptionLogger.Log(
                            $"DeepL returned HTTP {(int)response.StatusCode} ({response.StatusCode}). The response body was omitted because it may contain sensitive data.",
                            "DeepLService.TranslateAsync");
                        return null;
                    }

                    var result = JsonService.Deserialize<DeepLResponse>(response.Body);

                    if (result?.Translations != null && result.Translations.Count > 0)
                    {
                        return BoundedHttpResponseReader.ValidateTranslation(result.Translations[0].Text);
                    }
                }
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "DeepLService.TranslateAsync");
            }

            return null;
        }
    }
}
