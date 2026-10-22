using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace oPenEfficiency.Services
{
    public sealed class HtmlImportLimitException : IOException
    {
        public HtmlImportLimitException(string message) : base(message) { }
        public HtmlImportLimitException(string message, Exception innerException) : base(message, innerException) { }
    }

    public static class HtmlImportLimits
    {
        public const int MaxHtmlBytes = 5 * 1024 * 1024;
        public const int MaxStyleSheetBytes = 512 * 1024;
        public const int MaxCombinedStyleSheetBytes = 2 * 1024 * 1024;
        public const int MaxStyleSheetRequests = 16;
        public const int MaxMediaBytes = 10 * 1024 * 1024;
        public const int MaxCombinedMediaBytes = 50 * 1024 * 1024;
        public const int MaxMediaRequests = 32;
        public const int MaxRedirects = 3;
        public const int ResourceTimeoutMilliseconds = 10000;
        public static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);
    }

    public static class HtmlImportResourceReader
    {
        public static HttpWebRequest CreateBoundedRequest(Uri uri, int timeoutMilliseconds = HtmlImportLimits.ResourceTimeoutMilliseconds)
        {
            if (uri == null) throw new ArgumentNullException(nameof(uri));
            if (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)
                throw new ArgumentException("Only HTTP(S) import resources are supported.", nameof(uri));
            if (timeoutMilliseconds < 1) throw new ArgumentOutOfRangeException(nameof(timeoutMilliseconds));
            var request = (HttpWebRequest)WebRequest.Create(uri);
            request.AllowAutoRedirect = true;
            request.MaximumAutomaticRedirections = HtmlImportLimits.MaxRedirects;
            request.Timeout = timeoutMilliseconds;
            request.ReadWriteTimeout = timeoutMilliseconds;
            request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
            return request;
        }

        public static byte[] ReadBytes(Stream stream, int maxBytes, string limitMessage)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (maxBytes < 1) throw new ArgumentOutOfRangeException(nameof(maxBytes));
            using (var destination = new MemoryStream())
            {
                var buffer = new byte[8192];
                while (true)
                {
                    int read = stream.Read(buffer, 0, buffer.Length);
                    if (read == 0) break;
                    if (destination.Length > maxBytes - read)
                        throw new HtmlImportLimitException(limitMessage);
                    destination.Write(buffer, 0, read);
                }
                return destination.ToArray();
            }
        }
    }

    public sealed class HtmlImportDocument
    {
        public HtmlImportDocument() { Slides = new List<HtmlImportSlideModel>(); Warnings = new List<string>(); }
        public List<HtmlImportSlideModel> Slides { get; private set; }
        public List<string> Warnings { get; private set; }
    }

    public sealed class HtmlImportSlideModel
    {
        public int Number { get; set; }
        public string Id { get; set; }
        public bool IsGptDeck { get; set; }
        public bool IsGenericDom { get; set; }
        public string SectionOpenTag { get; set; }
        public string StyleSheet { get; set; }
        public List<string> Classes { get; set; }
        public bool IsDark { get; set; }
        public string RawHtml { get; set; }
        public string HeaderHtml { get; set; }
        public string ContentHtml { get; set; }
        public string FooterHtml { get; set; }
        public string TitleSlideHtml { get; set; }
        public string Title { get; set; }
        public string FooterLeft { get; set; }
        public string FooterRight { get; set; }
    }

    public sealed class HtmlImportDocumentPipeline
    {
        private readonly HtmlResourceLoader _resources;

        public HtmlImportDocumentPipeline(bool allowRemoteResources)
            : this(allowRemoteResources, HtmlImportLimits.ResourceTimeoutMilliseconds)
        {
        }

        public HtmlImportDocumentPipeline(bool allowRemoteResources, int resourceTimeoutMilliseconds)
        {
            _resources = new HtmlResourceLoader(allowRemoteResources, resourceTimeoutMilliseconds);
        }

        public HtmlImportDocument Load(string htmlFilePath)
        {
            if (string.IsNullOrWhiteSpace(htmlFilePath)) throw new ArgumentException("HTML file path is required.", "htmlFilePath");
            if (!File.Exists(htmlFilePath)) throw new FileNotFoundException("HTML file was not found.", htmlFilePath);
            string fullPath = Path.GetFullPath(htmlFilePath);
            return Parse(ReadHtmlFile(fullPath), Path.GetDirectoryName(fullPath));
        }

        public static string ReadHtmlFile(string htmlFilePath)
        {
            return ReadTextFile(htmlFilePath, HtmlImportLimits.MaxHtmlBytes, "The HTML file exceeds the 5 MiB import limit.");
        }

        public HtmlImportDocument Parse(string html, string baseDirectory)
        {
            try
            {
                var document = new HtmlImportDocument();
                if (string.IsNullOrWhiteSpace(html)) return document;
                string css = _resources.LoadCombinedStyleSheet(html, baseDirectory, document.Warnings);
                bool isGptDeck = html.IndexOf("class=\"deck\"", StringComparison.OrdinalIgnoreCase) >= 0 || css.IndexOf(".deck", StringComparison.OrdinalIgnoreCase) >= 0;
                var pattern = new Regex("<(?<tag>section|article)\\b(?<attrs>[^>]*)class\\s*=\\s*(?<quote>[\"'])(?<class>[^\"']*\\bslide\\b[^\"']*)\\k<quote>(?<after>[^>]*)>(?<content>.*?)</\\k<tag>\\s*>", RegexOptions.IgnoreCase | RegexOptions.Singleline, HtmlImportLimits.RegexTimeout);
                foreach (Match match in pattern.Matches(html))
                    AddSlide(document, match.Value.Substring(0, match.Value.IndexOf('>') + 1), match.Groups["content"].Value, css, isGptDeck);

                if (document.Slides.Count == 0)
                {
                    var fallback = new Regex("<(?<tag>section|article|div)\\b(?<attrs>[^>]*(?:data-slide(?:-index)?|role\\s*=\\s*[\"']slide[\"']|class\\s*=\\s*[\"'][^\"']*(?:deck-slide|presentation-slide)[^\"']*)[^>]*)>(?<content>.*?)</\\k<tag>\\s*>", RegexOptions.IgnoreCase | RegexOptions.Singleline, HtmlImportLimits.RegexTimeout);
                    foreach (Match match in fallback.Matches(html))
                        AddSlide(document, match.Value.Substring(0, match.Value.IndexOf('>') + 1), match.Groups["content"].Value, css, isGptDeck);
                    if (document.Slides.Count > 0) document.Warnings.Add("Presentation detected through alternative slide wrappers (data-slide, role=slide, or .deck-slide).");
                }
                AddDiagnostics(document, html, css);
                return document;
            }
            catch (RegexMatchTimeoutException ex)
            {
                throw new HtmlImportLimitException("The HTML import was stopped because its markup or styles exceeded the safe parsing-time limit.", ex);
            }
        }

        private static void AddSlide(HtmlImportDocument document, string openTag, string content, string css, bool isGptDeck)
        {
            var classes = SplitClasses(ReadAttribute(openTag, "class"));
            var model = new HtmlImportSlideModel
            {
                Number = document.Slides.Count + 1, Id = ReadAttribute(openTag, "id"), Classes = classes,
                SectionOpenTag = openTag, StyleSheet = css, IsGptDeck = isGptDeck, RawHtml = content ?? string.Empty
            };
            model.IsDark = classes.Any(c => string.Equals(c, "slide-dark", StringComparison.OrdinalIgnoreCase) || string.Equals(c, "dark", StringComparison.OrdinalIgnoreCase) || string.Equals(c, "hero", StringComparison.OrdinalIgnoreCase));
            model.Title = Clean(FirstTag(model.RawHtml, "h2"));
            if (string.IsNullOrWhiteSpace(model.Title)) model.Title = Clean(FirstTag(model.RawHtml, "h1"));
            model.HeaderHtml = BlockByClass(model.RawHtml, "div", "slide-header");
            model.ContentHtml = BlockByClass(model.RawHtml, "div", "slide-content");
            model.FooterHtml = BlockByClass(model.RawHtml, "div", "slide-footer");
            model.TitleSlideHtml = BlockByClass(model.RawHtml, "div", "title-slide-content");
            model.IsGenericDom = !isGptDeck && string.IsNullOrWhiteSpace(model.HeaderHtml) && string.IsNullOrWhiteSpace(model.ContentHtml) && string.IsNullOrWhiteSpace(model.TitleSlideHtml);
            var footer = Regex.Matches(model.FooterHtml ?? string.Empty, "<span\\b[^>]*>(?<value>.*?)</span>", RegexOptions.IgnoreCase | RegexOptions.Singleline).Cast<Match>().Select(m => Clean(m.Groups["value"].Value)).ToList();
            if (footer.Count > 0) model.FooterLeft = footer[0];
            if (footer.Count > 1) model.FooterRight = footer[1];
            document.Slides.Add(model);
        }

        private static void AddDiagnostics(HtmlImportDocument document, string html, string css)
        {
            if (document.Slides.Count == 0) document.Warnings.Add("No slide wrappers were detected. Supported wrappers are section/article.slide, data-slide, role=slide, and .deck-slide.");
            if (Regex.IsMatch(html, "<(canvas|iframe|video|audio)\\b", RegexOptions.IgnoreCase, HtmlImportLimits.RegexTimeout)) document.Warnings.Add("Interactive media is imported as a static placeholder or preview.");
            if (Regex.IsMatch(html, "<script\\b", RegexOptions.IgnoreCase, HtmlImportLimits.RegexTimeout)) document.Warnings.Add("JavaScript is not executed for security reasons.");
            if (Regex.IsMatch(css, "(animation|@keyframes|transition)\\s*[:{]", RegexOptions.IgnoreCase, HtmlImportLimits.RegexTimeout)) document.Warnings.Add("CSS animations are imported as a static initial state.");
            if (Regex.IsMatch(css, "(filter|backdrop-filter|mix-blend-mode)\\s*:", RegexOptions.IgnoreCase, HtmlImportLimits.RegexTimeout)) document.Warnings.Add("Filters, backdrop filters, and blend modes may render differently.");
            if (Regex.IsMatch(css, "@font-face", RegexOptions.IgnoreCase, HtmlImportLimits.RegexTimeout)) document.Warnings.Add("Web fonts are used when installed locally; otherwise PowerPoint uses a fallback font.");
        }

        private static string ReadAttribute(string tag, string name) { var m = Regex.Match(tag ?? "", "\\b" + Regex.Escape(name) + "\\s*=\\s*([\"'])(?<v>.*?)\\1", RegexOptions.IgnoreCase | RegexOptions.Singleline, HtmlImportLimits.RegexTimeout); return m.Success ? WebUtility.HtmlDecode(m.Groups["v"].Value) : null; }
        private static List<string> SplitClasses(string value) { return (value ?? "").Split((char[])null, StringSplitOptions.RemoveEmptyEntries).ToList(); }
        private static string FirstTag(string html, string tag) { var m = Regex.Match(html ?? "", "<" + tag + "\\b[^>]*>(?<v>.*?)</" + tag + ">", RegexOptions.IgnoreCase | RegexOptions.Singleline, HtmlImportLimits.RegexTimeout); return m.Success ? m.Groups["v"].Value : null; }
        private static string BlockByClass(string html, string tag, string name) { var m = Regex.Match(html ?? "", "<" + tag + "\\b[^>]*class\\s*=\\s*([\"'])[^\"']*\\b" + Regex.Escape(name) + "\\b[^\"']*\\1[^>]*>(?<v>.*?)</" + tag + ">", RegexOptions.IgnoreCase | RegexOptions.Singleline, HtmlImportLimits.RegexTimeout); return m.Success ? m.Groups["v"].Value : null; }
        private static string Clean(string html) { string value = Regex.Replace(html ?? "", "<br\\s*/?>", "\n", RegexOptions.IgnoreCase, HtmlImportLimits.RegexTimeout); value = Regex.Replace(value, "<.*?>", "", RegexOptions.Singleline, HtmlImportLimits.RegexTimeout); return Regex.Replace(WebUtility.HtmlDecode(value), "[ \\t]+", " ", RegexOptions.None, HtmlImportLimits.RegexTimeout).Trim(); }

        private static string ReadTextFile(string path, int maxBytes, string limitMessage)
        {
            var info = new FileInfo(path);
            if (info.Length > maxBytes) throw new HtmlImportLimitException(limitMessage);

            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 8192, FileOptions.SequentialScan))
            using (var destination = new MemoryStream())
            {
                var buffer = new byte[8192];
                while (true)
                {
                    int read = stream.Read(buffer, 0, buffer.Length);
                    if (read == 0) break;
                    if (destination.Length > maxBytes - read) throw new HtmlImportLimitException(limitMessage);
                    destination.Write(buffer, 0, read);
                }

                destination.Position = 0;
                using (var reader = new StreamReader(destination, Encoding.UTF8, true))
                    return reader.ReadToEnd();
            }
        }
    }

    public sealed class HtmlResourceLoader
    {
        private readonly bool _allowRemote;
        private readonly int _resourceTimeoutMilliseconds;
        private readonly Dictionary<string, string> _cache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private int _requestCount;
        private int _combinedCssBytes;
        public HtmlResourceLoader(bool allowRemote)
            : this(allowRemote, HtmlImportLimits.ResourceTimeoutMilliseconds)
        {
        }

        public HtmlResourceLoader(bool allowRemote, int resourceTimeoutMilliseconds)
        {
            if (resourceTimeoutMilliseconds < 1) throw new ArgumentOutOfRangeException(nameof(resourceTimeoutMilliseconds));
            _allowRemote = allowRemote;
            _resourceTimeoutMilliseconds = resourceTimeoutMilliseconds;
        }

        public string LoadCombinedStyleSheet(string html, string baseReference, IList<string> warnings)
        {
            var css = new StringBuilder();
            foreach (Match m in Regex.Matches(html ?? "", "<style\\b[^>]*>(?<v>.*?)</style>", RegexOptions.IgnoreCase | RegexOptions.Singleline, HtmlImportLimits.RegexTimeout)) AppendStyleSheet(css, ExpandImports(RebaseUrls(m.Groups["v"].Value, baseReference), baseReference, warnings, 0));
            foreach (Match m in Regex.Matches(html ?? "", "<link\\b[^>]*>", RegexOptions.IgnoreCase | RegexOptions.Singleline, HtmlImportLimits.RegexTimeout))
            {
                string tag = m.Value;
                if ((ReadAttribute(tag, "rel") ?? "").IndexOf("stylesheet", StringComparison.OrdinalIgnoreCase) >= 0) AppendStyleSheet(css, LoadExternal(ReadAttribute(tag, "href"), baseReference, warnings, 0));
            }
            return css.ToString();
        }

        public static string ResolveReference(string reference, string baseReference)
        {
            if (string.IsNullOrWhiteSpace(reference)) return null;
            string decoded = Uri.UnescapeDataString(WebUtility.HtmlDecode(reference.Trim())); Uri absolute;
            if (Uri.TryCreate(decoded, UriKind.Absolute, out absolute)) return absolute.IsFile ? absolute.LocalPath : absolute.AbsoluteUri;
            Uri baseUri;
            if (Uri.TryCreate(baseReference, UriKind.Absolute, out baseUri) && (baseUri.Scheme == "http" || baseUri.Scheme == "https")) return new Uri(baseUri.AbsoluteUri.EndsWith("/") ? baseUri : new Uri(baseUri, "."), decoded).AbsoluteUri;
            return Path.GetFullPath(Path.IsPathRooted(decoded) ? decoded : Path.Combine(baseReference ?? "", decoded.Replace('/', Path.DirectorySeparatorChar)));
        }

        private string LoadExternal(string href, string baseReference, IList<string> warnings, int depth)
        {
            if (depth > 4 || string.IsNullOrWhiteSpace(href)) return "";
            string resolved = ResolveReference(href, baseReference); string cached;
            if (string.IsNullOrWhiteSpace(resolved)) return "";
            if (_cache.TryGetValue(resolved, out cached)) return cached;
            _requestCount++;
            if (_requestCount > HtmlImportLimits.MaxStyleSheetRequests)
                throw new HtmlImportLimitException("The HTML import references more than 16 stylesheet resources.");
            Uri uri; string content;
            if (Uri.TryCreate(resolved, UriKind.Absolute, out uri) && (uri.Scheme == "http" || uri.Scheme == "https"))
            {
                if (!_allowRemote) { string warning = "Remote resource blocked: " + resolved; if (!warnings.Contains(warning)) warnings.Add(warning); return ""; }
                content = DownloadRemoteStyleSheet(uri);
            }
            else content = File.Exists(resolved)
                ? ReadBoundedStyleSheet(resolved)
                : "";
            EnsurePerResourceLimit(content);
            string nextBase = uri != null && (uri.Scheme == "http" || uri.Scheme == "https") ? new Uri(uri, ".").AbsoluteUri : Path.GetDirectoryName(resolved);
            return _cache[resolved] = ExpandImports(RebaseUrls(content, nextBase), nextBase, warnings, depth + 1);
        }

        private string ExpandImports(string css, string baseReference, IList<string> warnings, int depth) { if (depth > 4 || string.IsNullOrWhiteSpace(css)) return css ?? ""; return Regex.Replace(css, "@import\\s+(?:url\\(\\s*)?['\"]?(?<href>[^'\"\\s;)]+)['\"]?\\s*\\)?[^;]*;", m => LoadExternal(m.Groups["href"].Value, baseReference, warnings, depth + 1), RegexOptions.IgnoreCase, HtmlImportLimits.RegexTimeout); }
        private static string RebaseUrls(string css, string baseReference) { return Regex.Replace(css ?? "", "url\\(\\s*(['\"]?)(?<url>.*?)\\1\\s*\\)", m => { string value = m.Groups["url"].Value.Trim(); if (value.StartsWith("data:", StringComparison.OrdinalIgnoreCase) || value.StartsWith("#")) return m.Value; string resolved = ResolveReference(value, baseReference); return "url(\"" + resolved.Replace("\\", "/") + "\")"; }, RegexOptions.IgnoreCase | RegexOptions.Singleline, HtmlImportLimits.RegexTimeout); }
        private static string ReadAttribute(string tag, string name) { var m = Regex.Match(tag ?? "", "\\b" + Regex.Escape(name) + "\\s*=\\s*([\"'])(?<v>.*?)\\1", RegexOptions.IgnoreCase | RegexOptions.Singleline, HtmlImportLimits.RegexTimeout); return m.Success ? WebUtility.HtmlDecode(m.Groups["v"].Value) : null; }

        private void AppendStyleSheet(StringBuilder destination, string content)
        {
            int bytes = Encoding.UTF8.GetByteCount(content ?? string.Empty);
            if (_combinedCssBytes > HtmlImportLimits.MaxCombinedStyleSheetBytes - bytes)
                throw new HtmlImportLimitException("The combined stylesheets exceed the 2 MiB import limit.");
            _combinedCssBytes += bytes;
            destination.AppendLine(content);
        }

        private static void EnsurePerResourceLimit(string content)
        {
            if (Encoding.UTF8.GetByteCount(content ?? string.Empty) > HtmlImportLimits.MaxStyleSheetBytes)
                throw new HtmlImportLimitException("A stylesheet resource exceeds the 512 KiB import limit.");
        }

        private static string ReadBoundedStyleSheet(string path)
        {
            var info = new FileInfo(path);
            if (info.Length > HtmlImportLimits.MaxStyleSheetBytes)
                throw new HtmlImportLimitException("A stylesheet resource exceeds the 512 KiB import limit.");
            using (var reader = new StreamReader(new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 8192, FileOptions.SequentialScan), Encoding.UTF8, true))
            {
                var buffer = new char[8192];
                var result = new StringBuilder();
                int bytes = 0;
                while (true)
                {
                    int read = reader.Read(buffer, 0, buffer.Length);
                    if (read == 0) break;
                    bytes += Encoding.UTF8.GetByteCount(buffer, 0, read);
                    if (bytes > HtmlImportLimits.MaxStyleSheetBytes)
                        throw new HtmlImportLimitException("A stylesheet resource exceeds the 512 KiB import limit.");
                    result.Append(buffer, 0, read);
                }
                return result.ToString();
            }
        }

        private string DownloadRemoteStyleSheet(Uri uri)
        {
            var request = HtmlImportResourceReader.CreateBoundedRequest(uri, _resourceTimeoutMilliseconds);

            try
            {
                using (var response = (HttpWebResponse)request.GetResponse())
                {
                    if (uri.Scheme == Uri.UriSchemeHttps && response.ResponseUri.Scheme != Uri.UriSchemeHttps)
                        throw new HtmlImportLimitException("A remote stylesheet redirect attempted to downgrade HTTPS.");
                    if (response.ContentLength > HtmlImportLimits.MaxStyleSheetBytes)
                        throw new HtmlImportLimitException("A stylesheet resource exceeds the 512 KiB import limit.");

                    using (var stream = response.GetResponseStream())
                    using (var destination = new MemoryStream())
                    {
                        var buffer = new byte[8192];
                        while (true)
                        {
                            int read = stream.Read(buffer, 0, buffer.Length);
                            if (read == 0) break;
                            if (destination.Length > HtmlImportLimits.MaxStyleSheetBytes - read)
                                throw new HtmlImportLimitException("A stylesheet resource exceeds the 512 KiB import limit.");
                            destination.Write(buffer, 0, read);
                        }
                        return Encoding.UTF8.GetString(destination.ToArray());
                    }
                }
            }
            catch (WebException ex) when (ex.Status == WebExceptionStatus.Timeout)
            {
                throw new HtmlImportLimitException("A remote stylesheet exceeded the configured download timeout.", ex);
            }
        }
    }
}
