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

    public class LibreTranslateService
    {
        private readonly string _baseUrl;
        private readonly string _apiKey;
        private static readonly HttpClient _httpClient = new HttpClient();

        public LibreTranslateService(string baseUrl, string apiKey = null)
        {
            _baseUrl = string.IsNullOrWhiteSpace(baseUrl) ? "https://libretranslate.com/" : baseUrl;
            if (!_baseUrl.EndsWith("/")) _baseUrl += "/";
            _apiKey = apiKey;
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

                var requestData = new Dictionary<string, string>
                {
                    { "q", text },
                    { "source", "auto" },
                    { "target", target },
                    { "format", "text" }
                };

                if (!string.IsNullOrEmpty(_apiKey))
                {
                    requestData.Add("api_key", _apiKey);
                }

                string jsonContent = JsonService.Serialize(requestData);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                var response = await _httpClient.PostAsync(endpoint, content);
                string responseBody = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    System.Diagnostics.Debug.WriteLine($"LibreTranslate API Error ({response.StatusCode}): {responseBody}");
                    return null;
                }

                var result = JsonService.Deserialize<LibreTranslateResponse>(responseBody);
                if (!string.IsNullOrEmpty(result?.Error))
                {
                     System.Diagnostics.Debug.WriteLine($"LibreTranslate API Error: {result.Error}");
                     return null;
                }

                return result?.TranslatedText;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LibreTranslate Exception: {ex.Message}");
            }

            return null;
        }
    }
}
