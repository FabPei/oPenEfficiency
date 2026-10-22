using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services.OneNote
{
    /// <summary>
    /// Attach + safe access boundary for the desktop OneNote application object
    /// (FEAT-007 ON-S0). Replaces the former hierarchy-scan "first Page node"
    /// assumption: the active destination is resolved exclusively through
    /// Application.Windows.CurrentWindow.CurrentPageId, and every ambiguous or
    /// unavailable state fails closed with an explicit error.
    /// </summary>
    public sealed class OneNoteDesktopClient : IDisposable
    {
        private const string ProgId = "OneNote.Application";
        private const int XMLSchema2013 = 2;
        private const int PageInfoBasic = 0;

        private readonly dynamic _application;

        private OneNoteDesktopClient(dynamic application)
        {
            _application = application;
        }

        /// <summary>
        /// Connects to the desktop OneNote automation server, but only after a
        /// user-visible OneNote process has been found. Recent OneNote builds do
        /// not consistently publish OneNote.Application in the Running Object
        /// Table, so COM activation is the documented fallback and is guarded by
        /// the process check to ensure this command never starts OneNote itself.
        /// </summary>
        public static bool TryAttach(out OneNoteDesktopClient client, out string error)
        {
            client = null;
            error = null;
            if (!HasVisibleOneNoteWindow())
            {
                error = "OneNote Desktop is not running with a visible window.";
                return false;
            }

            try
            {
                object application;
                try
                {
                    application = Marshal.GetActiveObject(ProgId);
                }
                catch (COMException)
                {
                    Type applicationType = Type.GetTypeFromProgID(ProgId, throwOnError: false);
                    if (applicationType == null)
                    {
                        error = "The OneNote Desktop automation interface is not installed.";
                        return false;
                    }
                    application = Activator.CreateInstance(applicationType);
                }

                if (application == null)
                {
                    error = "The running OneNote Desktop instance could not be connected.";
                    return false;
                }
                client = new OneNoteDesktopClient(application);
                return true;
            }
            catch (COMException ex)
            {
                error = "The running OneNote Desktop instance did not expose its automation interface. No content was written.";
                ExceptionLogger.Log(ex, "OneNoteDesktopClient.TryAttach");
                return false;
            }
            catch (InvalidComObjectException ex)
            {
                error = "The running OneNote instance is no longer available.";
                ExceptionLogger.Log(ex, "OneNoteDesktopClient.TryAttach.InvalidComObject");
                return false;
            }
        }

        private static bool HasVisibleOneNoteWindow()
        {
            foreach (Process process in Process.GetProcessesByName("ONENOTE"))
            {
                using (process)
                {
                    try
                    {
                        if (process.MainWindowHandle != IntPtr.Zero) return true;
                    }
                    catch (InvalidOperationException)
                    {
                        // The process exited while it was being inspected.
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Resolves the user-visible active OneNote page through
        /// Windows.CurrentWindow.CurrentPageId. Fails closed when no window exists,
        /// the current window cannot be validated inside the Windows collection,
        /// or the page id is empty.
        /// </summary>
        public bool TryResolveActivePageId(out string pageId, out string error)
        {
            pageId = null;
            error = null;
            try
            {
                dynamic windows = _application.Windows;
                if (windows == null)
                {
                    error = "This OneNote Desktop build does not expose its open windows to Office automation. No content was written.";
                    return false;
                }
                if (windows.Count == 0)
                {
                    error = "No open OneNote window was found. Open (and focus) the destination page, then try again.";
                    return false;
                }

                dynamic current = null;
                try
                {
                    current = windows.CurrentWindow;
                }
                catch (COMException)
                {
                    current = null;
                }

                if (current == null)
                {
                    error = "OneNote's active window could not be identified. Focus the destination page, then try again.";
                    return false;
                }

                pageId = current.CurrentPageId as string;
                if (string.IsNullOrWhiteSpace(pageId))
                {
                    error = "The active OneNote window has no current page. Open (and focus) the destination page, then try again.";
                    return false;
                }

                return true;
            }
            catch (COMException ex)
            {
                error = "OneNote's active-window automation is unavailable. The active page could not be verified, so no content was written.";
                ExceptionLogger.Log(ex, "OneNoteDesktopClient.TryResolveActivePageId");
                return false;
            }
            catch (InvalidComObjectException ex)
            {
                error = "The running OneNote instance became unavailable. No content was written.";
                ExceptionLogger.Log(ex, "OneNoteDesktopClient.TryResolveActivePageId.InvalidComObject");
                return false;
            }
        }

        /// <summary>Reads one page's XML (basic page info, 2013 schema — unchanged from legacy).</summary>
        public bool TryGetPageXml(string pageId, out XDocument page, out string error)
        {
            page = null;
            error = null;
            try
            {
                string pageXml;
                _application.GetPageContent(pageId, out pageXml, PageInfoBasic, XMLSchema2013);
                if (string.IsNullOrEmpty(pageXml))
                {
                    error = "The OneNote page content could not be read. It may have been moved, renamed, or deleted.";
                    return false;
                }
                page = XDocument.Parse(pageXml, LoadOptions.PreserveWhitespace);
                string returnedPageId = page.Root?.Attribute("ID")?.Value;
                if (string.IsNullOrWhiteSpace(returnedPageId) ||
                    !string.Equals(returnedPageId, pageId, StringComparison.OrdinalIgnoreCase))
                {
                    page = null;
                    error = "OneNote returned content for an unexpected page. The update was cancelled.";
                    return false;
                }
                return true;
            }
            catch (COMException ex)
            {
                error = "The OneNote page content could not be read. It may have been moved, renamed, or deleted.";
                ExceptionLogger.Log(ex, "OneNoteDesktopClient.TryGetPageXml");
                return false;
            }
        }

        /// <summary>
        /// Writes page XML with expected-last-modified concurrency protection.
        /// force is always false: conflicts must surface as explicit errors instead of
        /// silently overwriting concurrent page edits.
        /// </summary>
        public bool TryUpdatePage(XDocument page, DateTime expectedLastModified, out string error)
        {
            error = null;
            if (page?.Root == null)
            {
                error = "The OneNote page update was empty. No content was written.";
                return false;
            }
            if (expectedLastModified == DateTime.MinValue)
            {
                error = "OneNote did not provide a valid page version. The update was cancelled to avoid overwriting newer changes.";
                return false;
            }

            try
            {
                _application.UpdatePageContent(
                    page.ToString(SaveOptions.DisableFormatting),
                    expectedLastModified,
                    XMLSchema2013,
                    false);
                return true;
            }
            catch (COMException ex)
            {
                error = "The OneNote page changed while the label was being written. The update was skipped; run Refresh again.";
                ExceptionLogger.Log(ex, "OneNoteDesktopClient.TryUpdatePage");
                return false;
            }
        }

        public void Dispose()
        {
            object application = _application;
            if (application != null && Marshal.IsComObject(application))
            {
                try
                {
                    Marshal.ReleaseComObject(application);
                }
                catch (InvalidComObjectException)
                {
                    // The OneNote process already released the automation object.
                }
            }
        }
    }
}
