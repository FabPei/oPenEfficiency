using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Security.Principal;
using System.Xml.Linq;
using oPenEfficiency.OneNote;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services.OneNote
{
    /// <summary>PowerPoint-side client for the current-user OneNote in-process bridge.</summary>
    public sealed class OneNoteBridgeClient
    {
        private const int ConnectTimeoutMilliseconds = 1500;
        private const int OperationTimeoutMilliseconds = 10000;
        private readonly string _pipeName;

        public OneNoteBridgeClient()
            : this(OneNoteBridgeProtocol.GetPipeNameForCurrentUser())
        {
        }

        internal OneNoteBridgeClient(string pipeName)
        {
            _pipeName = pipeName ?? throw new ArgumentNullException(nameof(pipeName));
        }

        public bool TryGetActivePageId(out string pageId, out string error)
        {
            pageId = null;
            if (!TryExecute(new OneNoteBridgeRequest { Operation = OneNoteBridgeProtocol.GetActivePage },
                    out OneNoteBridgeResponse response, out error)) return false;
            pageId = response.PageId;
            if (string.IsNullOrWhiteSpace(pageId))
            {
                error = "The OneNote helper did not return an active page. No content was written.";
                return false;
            }
            return true;
        }

        public bool TryReadPage(string pageId, out XDocument page, out DateTime lastModified, out string error)
        {
            page = null;
            lastModified = DateTime.MinValue;
            if (!TryExecute(new OneNoteBridgeRequest
                {
                    Operation = OneNoteBridgeProtocol.ReadPage,
                    PageId = pageId
                }, out OneNoteBridgeResponse response, out error)) return false;

            try
            {
                page = XDocument.Parse(response.PageXml, LoadOptions.PreserveWhitespace);
                lastModified = ConnectedTextPayload.ReadLastModified(page);
                if (lastModified == DateTime.MinValue)
                {
                    page = null;
                    error = "The OneNote helper returned a page without a valid version.";
                    return false;
                }
                return true;
            }
            catch (Exception ex) when (ex is ArgumentNullException || ex is System.Xml.XmlException)
            {
                ExceptionLogger.Log(ex, "OneNoteBridgeClient.TryReadPage");
                error = "The OneNote helper returned invalid page content.";
                return false;
            }
        }

        public bool TryAppendConnectedText(string pageId, string sourceId, string text, out string error) =>
            TryExecute(new OneNoteBridgeRequest
            {
                Operation = OneNoteBridgeProtocol.AppendConnectedText,
                PageId = pageId,
                SourceId = sourceId,
                Text = text
            }, out _, out error);

        public bool TryRefreshConnectedTexts(
            string pageId,
            IReadOnlyList<OneNoteBridgeTarget> targets,
            out int replacedCount,
            out string error)
        {
            replacedCount = 0;
            if (!TryExecute(new OneNoteBridgeRequest
                {
                    Operation = OneNoteBridgeProtocol.RefreshConnectedTexts,
                    PageId = pageId,
                    Targets = targets == null ? null : new List<OneNoteBridgeTarget>(targets)
                }, out OneNoteBridgeResponse response, out error)) return false;
            replacedCount = response.ReplacedCount;
            return true;
        }

        private bool TryExecute(OneNoteBridgeRequest request, out OneNoteBridgeResponse response, out string error)
        {
            response = null;
            error = null;
            try
            {
                using (var pipe = new NamedPipeClientStream(
                    ".",
                    _pipeName,
                    PipeDirection.InOut,
                    PipeOptions.None,
                    TokenImpersonationLevel.Identification))
                {
                    pipe.Connect(ConnectTimeoutMilliseconds);
                    pipe.ReadTimeout = OperationTimeoutMilliseconds;
                    pipe.WriteTimeout = OperationTimeoutMilliseconds;
                    OneNoteBridgeProtocol.WriteRequest(pipe, request);
                    response = OneNoteBridgeProtocol.ReadResponse(pipe);
                }

                if (response == null || response.Version != OneNoteBridgeProtocol.Version)
                {
                    error = "The PowerPoint and OneNote helper versions do not match. Restart both applications after updating openEfficiency.";
                    return false;
                }
                if (!response.Success)
                {
                    error = response.Error ?? "The OneNote helper rejected the request. No content was written.";
                    return false;
                }
                return true;
            }
            catch (TimeoutException ex)
            {
                ExceptionLogger.Log(ex, "OneNoteBridgeClient.Connect");
                error = "The openEfficiency OneNote helper is not running. Ensure the OneNote add-in is enabled, then restart OneNote and try again. No content was written.";
                return false;
            }
            catch (IOException ex)
            {
                ExceptionLogger.Log(ex, "OneNoteBridgeClient.Transport");
                error = "The openEfficiency OneNote helper did not respond. Restart OneNote and try again. No content was written.";
                return false;
            }
            catch (UnauthorizedAccessException ex)
            {
                ExceptionLogger.Log(ex, "OneNoteBridgeClient.Access");
                error = "Access to the current user's OneNote helper was denied. No content was written.";
                return false;
            }
        }
    }
}
