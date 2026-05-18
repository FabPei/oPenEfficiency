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
        private static readonly HttpClient _httpClient = new HttpClient();

        public DeepLService(string apiKey)
        {
            _apiKey = apiKey;
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

                var response = await _httpClient.PostAsync(endpoint, content);
                if (!response.IsSuccessStatusCode)
                {
                    string error = await response.Content.ReadAsStringAsync();
                    System.Diagnostics.Debug.WriteLine($"DeepL API Error ({response.StatusCode}): {error}");
                    return null;
                }

                string jsonResponse = await response.Content.ReadAsStringAsync();
                var result = JsonService.Deserialize<DeepLResponse>(jsonResponse);

                if (result?.Translations != null && result.Translations.Count > 0)
                {
                    return result.Translations[0].Text;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DeepL Translation Exception: {ex.Message}");
            }

            return null;
        }
    }
}
