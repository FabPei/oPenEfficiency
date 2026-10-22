using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace oPenEfficiency.Services
{
    internal sealed class BoundedHttpResponse
    {
        internal BoundedHttpResponse(HttpStatusCode statusCode, string body)
        {
            StatusCode = statusCode;
            Body = body;
        }

        internal HttpStatusCode StatusCode { get; }
        internal string Body { get; }
        internal bool IsSuccessStatusCode => (int)StatusCode >= 200 && (int)StatusCode <= 299;
    }

    internal static class BoundedHttpResponseReader
    {
        internal const int MaxResponseBytes = 1024 * 1024;
        internal const int MaxTranslationCharacters = 200000;
        internal static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(30);

        internal static async Task<BoundedHttpResponse> SendAsync(
            HttpClient httpClient,
            HttpRequestMessage request,
            TimeSpan timeout)
        {
            if (httpClient == null) throw new ArgumentNullException(nameof(httpClient));
            if (request == null) throw new ArgumentNullException(nameof(request));
            if (timeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(timeout));

            using (var timeoutSource = new CancellationTokenSource(timeout))
            {
                try
                {
                    using (var response = await httpClient.SendAsync(
                        request,
                        HttpCompletionOption.ResponseHeadersRead,
                        timeoutSource.Token).ConfigureAwait(false))
                    {
                        if (!response.IsSuccessStatusCode)
                            return new BoundedHttpResponse(response.StatusCode, null);

                        string body = await ReadBodyAsync(response.Content, timeoutSource.Token).ConfigureAwait(false);
                        return new BoundedHttpResponse(response.StatusCode, body);
                    }
                }
                catch (OperationCanceledException ex) when (timeoutSource.IsCancellationRequested)
                {
                    throw new TimeoutException("The translation provider response exceeded the request timeout.", ex);
                }
            }
        }

        internal static string ValidateTranslation(string translation)
        {
            if (translation != null && translation.Length > MaxTranslationCharacters)
                throw new InvalidDataException("The translation provider output exceeded the allowed character limit.");

            return translation;
        }

        private static async Task<string> ReadBodyAsync(HttpContent content, CancellationToken cancellationToken)
        {
            if (content == null)
                return string.Empty;

            long? declaredLength = content.Headers.ContentLength;
            if (declaredLength.HasValue && declaredLength.Value > MaxResponseBytes)
                throw new InvalidDataException("The translation provider response exceeded the allowed byte limit.");

            using (var source = await content.ReadAsStreamAsync().ConfigureAwait(false))
            using (var destination = new MemoryStream())
            {
                var buffer = new byte[8192];
                int total = 0;
                while (true)
                {
                    int read = await source.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
                    if (read == 0)
                        break;

                    if (total > MaxResponseBytes - read)
                        throw new InvalidDataException("The translation provider response exceeded the allowed byte limit.");

                    destination.Write(buffer, 0, read);
                    total += read;
                }

                if (declaredLength.HasValue && declaredLength.Value != total)
                    throw new InvalidDataException("The translation provider response was truncated.");

                Encoding encoding = GetEncoding(content);
                return encoding.GetString(destination.GetBuffer(), 0, total);
            }
        }

        private static Encoding GetEncoding(HttpContent content)
        {
            string charset = content.Headers.ContentType?.CharSet?.Trim('"');
            if (string.IsNullOrWhiteSpace(charset))
                return new UTF8Encoding(false, true);

            try
            {
                return Encoding.GetEncoding(
                    charset,
                    EncoderFallback.ExceptionFallback,
                    DecoderFallback.ExceptionFallback);
            }
            catch (ArgumentException ex)
            {
                throw new InvalidDataException("The translation provider response declared an unsupported character encoding.", ex);
            }
        }
    }
}
