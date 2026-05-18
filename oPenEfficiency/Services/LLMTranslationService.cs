using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace oPenEfficiency.Services
{
    #region OpenAI Models
    [DataContract]
    public class OpenAIRequest
    {
        [DataMember(Name = "model")]
        public string Model { get; set; }
        [DataMember(Name = "messages")]
        public List<OpenAIMessage> Messages { get; set; }
    }

    [DataContract]
    public class OpenAIMessage
    {
        [DataMember(Name = "role")]
        public string Role { get; set; }
        [DataMember(Name = "content")]
        public string Content { get; set; }
    }

    [DataContract]
    public class OpenAIResponse
    {
        [DataMember(Name = "choices")]
        public List<OpenAIChoice> Choices { get; set; }
    }

    [DataContract]
    public class OpenAIChoice
    {
        [DataMember(Name = "message")]
        public OpenAIMessage Message { get; set; }
    }
    #endregion

    #region Claude Models
    [DataContract]
    public class ClaudeRequest
    {
        [DataMember(Name = "model")]
        public string Model { get; set; }
        [DataMember(Name = "max_tokens")]
        public int MaxTokens { get; set; } = 1024;
        [DataMember(Name = "messages")]
        public List<ClaudeMessage> Messages { get; set; }
    }

    [DataContract]
    public class ClaudeMessage
    {
        [DataMember(Name = "role")]
        public string Role { get; set; }
        [DataMember(Name = "content")]
        public string Content { get; set; }
    }

    [DataContract]
    public class ClaudeResponse
    {
        [DataMember(Name = "content")]
        public List<ClaudeContent> Content { get; set; }
    }

    [DataContract]
    public class ClaudeContent
    {
        [DataMember(Name = "text")]
        public string Text { get; set; }
    }
    #endregion

    public class LLMTranslationService
    {
        private static readonly HttpClient _httpClient = new HttpClient();

        public async Task<string> TranslateWithOpenAIAsync(string text, string targetLanguage, string apiKey, string model)
        {
            if (string.IsNullOrEmpty(apiKey)) return null;

            try
            {
                // Force TLS 1.2 for modern API compatibility in .NET 4.8
                System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;

                var requestBody = new OpenAIRequest
                {
                    Model = model,
                    Messages = new List<OpenAIMessage>
                    {
                        new OpenAIMessage { Role = "system", Content = $"You are a professional translator. Translate the following text to {targetLanguage}. Output ONLY the translation, no extra text or explanations." },
                        new OpenAIMessage { Role = "user", Content = text }
                    }
                };

                string jsonRequest = JsonService.Serialize(requestBody);
                var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");
                
                using (var request = new HttpRequestMessage(HttpMethod.Post, "https://api.openai.com/v1/chat/completions"))
                {
                    request.Headers.Add("Authorization", $"Bearer {apiKey}");
                    request.Content = content;
                    
                    var response = await _httpClient.SendAsync(request);
                    if (!response.IsSuccessStatusCode)
                    {
                        string error = await response.Content.ReadAsStringAsync();
                        System.Diagnostics.Debug.WriteLine($"OpenAI API Error ({response.StatusCode}): {error}");
                        return null;
                    }

                    string jsonResponse = await response.Content.ReadAsStringAsync();
                    var result = JsonService.Deserialize<OpenAIResponse>(jsonResponse);

                    return result?.Choices?[0]?.Message?.Content?.Trim();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OpenAI Translation Exception: {ex.Message}");
                return null;
            }
        }

        public async Task<string> TranslateWithClaudeAsync(string text, string targetLanguage, string apiKey, string model)
        {
            if (string.IsNullOrEmpty(apiKey)) return null;

            try
            {
                // Force TLS 1.2 for modern API compatibility in .NET 4.8
                System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;

                var requestBody = new ClaudeRequest
                {
                    Model = model,
                    Messages = new List<ClaudeMessage>
                    {
                        new ClaudeMessage { Role = "user", Content = $"Translate the following text to {targetLanguage}. Output ONLY the translation, no extra text.\n\nText: {text}" }
                    }
                };

                string jsonRequest = JsonService.Serialize(requestBody);
                var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");
                
                using (var request = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages"))
                {
                    request.Headers.Add("x-api-key", apiKey);
                    request.Headers.Add("anthropic-version", "2023-06-01");
                    request.Content = content;
                    
                    var response = await _httpClient.SendAsync(request);
                    if (!response.IsSuccessStatusCode)
                    {
                        string error = await response.Content.ReadAsStringAsync();
                        System.Diagnostics.Debug.WriteLine($"Claude API Error ({response.StatusCode}): {error}");
                        return null;
                    }

                    string jsonResponse = await response.Content.ReadAsStringAsync();
                    var result = JsonService.Deserialize<ClaudeResponse>(jsonResponse);

                    return result?.Content?[0]?.Text?.Trim();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Claude Translation Exception: {ex.Message}");
                return null;
            }
        }
    }
}
