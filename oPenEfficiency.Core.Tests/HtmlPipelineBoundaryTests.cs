using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class HtmlPipelineBoundaryTests
    {
        private string _directory;

        [TestInitialize]
        public void SetUp()
        {
            _directory = Path.Combine(Path.GetTempPath(), "oPE-html", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_directory);
        }

        [TestCleanup]
        public void TearDown()
        {
            if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
        }

        [TestMethod]
        public void HtmlPipelineExtractsDecodedMetadataAndDarkMode()
        {
            string html = "<section class='slide hero custom' id='a&amp;b'>" +
                "<h2>Hello &amp;<br>World</h2>" +
                "<div class='slide-footer'><span>Left &lt;x&gt;</span><span>Right</span></div>" +
                "</section>";

            HtmlImportSlideModel slide = new HtmlImportDocumentPipeline(false)
                .Parse(html, _directory)
                .Slides.Single();

            Assert.AreEqual("a&b", slide.Id);
            Assert.AreEqual("Hello &\nWorld", slide.Title);
            Assert.AreEqual("Left <x>", slide.FooterLeft);
            Assert.AreEqual("Right", slide.FooterRight);
            Assert.IsTrue(slide.IsDark);
            CollectionAssert.AreEqual(new[] { "slide", "hero", "custom" }, slide.Classes);
        }

        [TestMethod]
        public void HtmlPipelineReportsUnsupportedDynamicContent()
        {
            string html = "<style>@font-face{} .x{animation:go 1s;filter:blur(1px)}</style>" +
                "<script>alert(1)</script><section class='slide'><canvas></canvas><video></video></section>";

            HtmlImportDocument document = new HtmlImportDocumentPipeline(false).Parse(html, _directory);

            Assert.IsTrue(document.Warnings.Any(warning => warning.Contains("JavaScript")));
            Assert.IsTrue(document.Warnings.Any(warning => warning.Contains("Interactive media")));
            Assert.IsTrue(document.Warnings.Any(warning => warning.Contains("animations")));
            Assert.IsTrue(document.Warnings.Any(warning => warning.Contains("Filters")));
            Assert.IsTrue(document.Warnings.Any(warning => warning.Contains("Web fonts")));
        }

        [TestMethod]
        public void HtmlPipelineBlocksRemoteStyleSheetsWithoutNetworkAccess()
        {
            string html = "<link rel='stylesheet' href='https://example.invalid/theme.css'>" +
                "<section class='slide'><h1>One</h1></section>";

            HtmlImportDocument document = new HtmlImportDocumentPipeline(false).Parse(html, _directory);

            Assert.AreEqual(1, document.Slides.Count);
            Assert.IsTrue(document.Warnings.Single().Contains("Remote resource blocked"));
            Assert.AreEqual(string.Empty, document.Slides[0].StyleSheet.Trim());
        }

        [TestMethod]
        public void HtmlPipelineRejectsOversizedHtmlBeforeParsing()
        {
            string path = Path.Combine(_directory, "oversized.html");
            using (var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
            {
                stream.SetLength(HtmlImportLimits.MaxHtmlBytes + 1L);
            }

            HtmlImportLimitException error = Assert.ThrowsException<HtmlImportLimitException>(
                () => new HtmlImportDocumentPipeline(false).Load(path));

            StringAssert.Contains(error.Message, "5 MiB");
        }

        [TestMethod]
        public void HtmlPipelineRejectsOversizedLocalStyleSheet()
        {
            string cssPath = Path.Combine(_directory, "oversized.css");
            using (var stream = new FileStream(cssPath, FileMode.CreateNew, FileAccess.Write))
            {
                stream.SetLength(HtmlImportLimits.MaxStyleSheetBytes + 1L);
            }
            string html = "<link rel='stylesheet' href='oversized.css'><section class='slide'></section>";

            HtmlImportLimitException error = Assert.ThrowsException<HtmlImportLimitException>(
                () => new HtmlImportDocumentPipeline(false).Parse(html, _directory));

            StringAssert.Contains(error.Message, "512 KiB");
        }

        [TestMethod]
        public void HtmlPipelineRejectsTooManyStyleSheetResources()
        {
            var html = new System.Text.StringBuilder();
            for (int index = 0; index <= HtmlImportLimits.MaxStyleSheetRequests; index++)
            {
                string name = "style-" + index + ".css";
                File.WriteAllText(Path.Combine(_directory, name), ".x{color:red}");
                html.Append("<link rel='stylesheet' href='").Append(name).Append("'>");
            }
            html.Append("<section class='slide'></section>");

            HtmlImportLimitException error = Assert.ThrowsException<HtmlImportLimitException>(
                () => new HtmlImportDocumentPipeline(false).Parse(html.ToString(), _directory));

            StringAssert.Contains(error.Message, "more than 16");
        }

        [TestMethod]
        public void HtmlPipelineRejectsExcessiveCombinedCss()
        {
            var html = new System.Text.StringBuilder();
            string css = new string('x', HtmlImportLimits.MaxStyleSheetBytes);
            for (int index = 0; index < 5; index++)
            {
                string name = "large-" + index + ".css";
                File.WriteAllText(Path.Combine(_directory, name), css);
                html.Append("<link rel='stylesheet' href='").Append(name).Append("'>");
            }
            html.Append("<section class='slide'></section>");

            HtmlImportLimitException error = Assert.ThrowsException<HtmlImportLimitException>(
                () => new HtmlImportDocumentPipeline(false).Parse(html.ToString(), _directory));

            StringAssert.Contains(error.Message, "combined stylesheets");
        }

        [TestMethod]
        public void HtmlResourceReaderStopsUnknownLengthMediaAtTheByteLimit()
        {
            using (var stream = new MemoryStream(new byte[HtmlImportLimits.MaxMediaBytes + 1]))
            {
                HtmlImportLimitException error = Assert.ThrowsException<HtmlImportLimitException>(
                    () => HtmlImportResourceReader.ReadBytes(stream, HtmlImportLimits.MaxMediaBytes, "media limit"));
                Assert.AreEqual("media limit", error.Message);
                Assert.IsTrue(stream.Position <= HtmlImportLimits.MaxMediaBytes + 1L);
            }
        }

        [TestMethod]
        public void HtmlRemoteRequestsHaveBoundedTimeoutsAndRedirects()
        {
            var request = HtmlImportResourceReader.CreateBoundedRequest(new Uri("https://example.invalid/resource.css"));

            Assert.AreEqual(HtmlImportLimits.ResourceTimeoutMilliseconds, request.Timeout);
            Assert.AreEqual(HtmlImportLimits.ResourceTimeoutMilliseconds, request.ReadWriteTimeout);
            Assert.AreEqual(HtmlImportLimits.MaxRedirects, request.MaximumAutomaticRedirections);
            Assert.IsTrue(request.AllowAutoRedirect);
            Assert.IsTrue(HtmlImportLimits.RegexTimeout <= TimeSpan.FromSeconds(2));
        }

        [TestMethod]
        public void HtmlPipelineStopsSlowRemoteStyleSheetAtConfiguredTimeout()
        {
            using (var server = new SlowLoopbackServer(TimeSpan.FromMilliseconds(750)))
            {
                string html = "<link rel='stylesheet' href='" + server.ResourceUri.AbsoluteUri +
                    "'><section class='slide'></section>";
                var stopwatch = Stopwatch.StartNew();

                HtmlImportLimitException error = Assert.ThrowsException<HtmlImportLimitException>(
                    () => new HtmlImportDocumentPipeline(true, 100).Parse(html, _directory));

                stopwatch.Stop();
                StringAssert.Contains(error.Message, "timeout");
                Assert.IsTrue(stopwatch.Elapsed < TimeSpan.FromSeconds(3),
                    "The slow resource was not interrupted promptly: " + stopwatch.Elapsed);
            }
        }

        [TestMethod]
        public void HtmlPipelineBoundsPathologicalMalformedMarkup()
        {
            string fixturePath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Fixtures",
                "html",
                "pathological-unclosed-slide.html");
            string fragment = File.ReadAllText(fixturePath);
            string html = string.Concat(Enumerable.Repeat(fragment, 20000));
            var stopwatch = Stopwatch.StartNew();

            try
            {
                HtmlImportDocument document = new HtmlImportDocumentPipeline(false).Parse(html, _directory);
                Assert.AreEqual(0, document.Slides.Count);
            }
            catch (HtmlImportLimitException error)
            {
                StringAssert.Contains(error.Message, "safe parsing-time limit");
                Assert.IsInstanceOfType(error.InnerException, typeof(System.Text.RegularExpressions.RegexMatchTimeoutException));
            }

            stopwatch.Stop();
            Assert.IsTrue(stopwatch.Elapsed < TimeSpan.FromSeconds(7),
                "Pathological markup exceeded the bounded parsing window: " + stopwatch.Elapsed);
        }

        [TestMethod]
        public void HtmlPipelineStopsCssImportExpansionAtTheDepthLimit()
        {
            for (int index = 0; index < 7; index++)
            {
                string next = index == 6 ? ".deep{color:red}" : "@import 'depth-" + (index + 1) + ".css';";
                File.WriteAllText(Path.Combine(_directory, "depth-" + index + ".css"), next);
            }
            string html = "<link rel='stylesheet' href='depth-0.css'><section class='slide'></section>";

            string css = new HtmlImportDocumentPipeline(false).Parse(html, _directory).Slides.Single().StyleSheet;

            Assert.IsFalse(css.Contains(".deep"));
        }

        [TestMethod]
        public void HtmlPipelineExpandsNestedLocalCssAndRebasesUrls()
        {
            Directory.CreateDirectory(Path.Combine(_directory, "css", "nested"));
            File.WriteAllText(
                Path.Combine(_directory, "css", "main.css"),
                "@import 'nested/colors.css'; .hero{background:url('../images/bg.png')}");
            File.WriteAllText(
                Path.Combine(_directory, "css", "nested", "colors.css"),
                ".title{color:red;background:url('../../icons/a.svg')}");
            string html = "<link rel='stylesheet' href='css/main.css'><section class='slide'><h1>One</h1></section>";

            string css = new HtmlImportDocumentPipeline(false).Parse(html, _directory).Slides[0].StyleSheet;

            StringAssert.Contains(css, ".title{color:red");
            StringAssert.Contains(css, Path.Combine(_directory, "icons", "a.svg").Replace("\\", "/"));
            StringAssert.Contains(css, Path.Combine(_directory, "images", "bg.png").Replace("\\", "/"));
            Assert.IsFalse(css.Contains("@import"));
        }

        [TestMethod]
        public void AlternativeSlideWrappersAreDetectedInDocumentOrder()
        {
            string html = "<article role='slide'><h1>First</h1></article>" +
                "<div data-slide-index='2'><h1>Second</h1></div>";

            HtmlImportDocument document = new HtmlImportDocumentPipeline(false).Parse(html, _directory);

            CollectionAssert.AreEqual(new[] { "First", "Second" }, document.Slides.Select(slide => slide.Title).ToArray());
            CollectionAssert.AreEqual(new[] { 1, 2 }, document.Slides.Select(slide => slide.Number).ToArray());
            Assert.IsTrue(document.Warnings.Any(warning => warning.Contains("alternative slide wrappers")));
        }

        [TestMethod]
        public void ResourceResolutionHandlesFilesUrlsEntitiesAndDataUris()
        {
            Assert.AreEqual(
                Path.Combine(_directory, "images", "a b.svg"),
                HtmlResourceLoader.ResolveReference("images/a%20b.svg", _directory));
            Assert.AreEqual(
                "https://example.com/base/image.png",
                HtmlResourceLoader.ResolveReference("image.png", "https://example.com/base/index.html"));
            Assert.AreEqual(
                "https://example.com/a.css?x=1&y=2",
                HtmlResourceLoader.ResolveReference("https://example.com/a.css?x=1&amp;y=2", _directory));
        }

        [TestMethod]
        public void HtmlLayoutFitPreservesAspectRatioAndCentersContent()
        {
            HtmlSlideScale scale = new HtmlLayoutEngine().Fit(100, 50, 300, 300);

            Assert.AreEqual(3, scale.X, 1e-6);
            Assert.AreEqual(3, scale.Y, 1e-6);
            Assert.AreEqual(0, scale.OffsetX, 1e-6);
            Assert.AreEqual(75, scale.OffsetY, 1e-6);
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new HtmlLayoutEngine().Fit(0, 50, 100, 100));
        }

        [TestMethod]
        public void HtmlGridAndColumnsNeverReturnNegativeCellSizes()
        {
            var engine = new HtmlLayoutEngine();
            var columns = engine.Columns(new HtmlLayoutRect(0, 0, 10, 20), 3, 20);
            var grid = engine.Grid(new HtmlLayoutRect(0, 0, 10, 10), 5, 2, 20, 20);

            Assert.IsTrue(columns.All(cell => cell.Width >= 0 && cell.Height >= 0));
            Assert.IsTrue(grid.All(cell => cell.Width >= 0 && cell.Height >= 0));
            Assert.AreEqual(0, engine.Columns(new HtmlLayoutRect(0, 0, 10, 10), 0, 0).Count);
            Assert.AreEqual(0, engine.Grid(new HtmlLayoutRect(0, 0, 10, 10), 2, 0, 0, 0).Count);
        }

        private sealed class SlowLoopbackServer : IDisposable
        {
            private readonly TcpListener _listener;
            private readonly TimeSpan _delay;
            private readonly Task _serverTask;

            public SlowLoopbackServer(TimeSpan delay)
            {
                _delay = delay;
                _listener = new TcpListener(IPAddress.Loopback, 0);
                _listener.Start();
                int port = ((IPEndPoint)_listener.LocalEndpoint).Port;
                ResourceUri = new Uri("http://127.0.0.1:" + port + "/slow.css");
                _serverTask = Task.Run((Action)ServeOneRequest);
            }

            public Uri ResourceUri { get; private set; }

            public void Dispose()
            {
                _listener.Stop();
                try { _serverTask.Wait(TimeSpan.FromSeconds(2)); }
                catch (AggregateException) { }
            }

            private void ServeOneRequest()
            {
                try
                {
                    using (TcpClient client = _listener.AcceptTcpClient())
                    using (NetworkStream stream = client.GetStream())
                    {
                        ReadRequestHeaders(stream);
                        Thread.Sleep(_delay);
                        byte[] response = Encoding.ASCII.GetBytes(
                            "HTTP/1.1 200 OK\r\nContent-Type: text/css\r\nContent-Length: 13\r\nConnection: close\r\n\r\n.x{color:red}");
                        stream.Write(response, 0, response.Length);
                    }
                }
                catch (IOException) { }
                catch (ObjectDisposedException) { }
                catch (SocketException) { }
            }

            private static void ReadRequestHeaders(Stream stream)
            {
                int matched = 0;
                byte[] terminator = { 13, 10, 13, 10 };
                while (matched < terminator.Length)
                {
                    int value = stream.ReadByte();
                    if (value < 0) return;
                    matched = value == terminator[matched] ? matched + 1 : (value == terminator[0] ? 1 : 0);
                }
            }
        }
    }
}
