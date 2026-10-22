using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    [ComVisible(true)]
    public class MainRibbon : IRibbonExtensibility
    {
        private IRibbonUI ribbon;

        // Cache the converted picture so Office's loadImage callback loads the icon only once.
        private object _ribbonIcon;

        public MainRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("oPenEfficiency.UI.MainRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        /// <summary>
        /// Ribbon build hint: labels the loaded assembly's write time, so a running session's
        /// build is visible without opening the log (see StartupBuildLog).
        /// </summary>
        public string GetBuildStampLabel(IRibbonControl control)
        {
            try
            {
                return StartupBuildLog.FormatBuildHint(typeof(ThisAddIn).Assembly);
            }
            catch (Exception ex)
            {
                // The ribbon must render even when the build hint cannot be read.
                ExceptionLogger.Log(ex, "MainRibbon.GetBuildStampLabel");
                return string.Empty;
            }
        }

        /// <summary>
        /// Dev gate for the QA Status ribbon button: hidden whenever the repository checkout is
        /// not reachable (FeatureAcceptanceStatusService.IsAvailable is false in production).
        /// </summary>
        public bool GetQaWindowVisible(IRibbonControl control)
        {
            try
            {
                return Services.FeatureAcceptanceStatusService.IsAvailable;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MainRibbon.GetQaWindowVisible");
                return false;
            }
        }

        public void OnOpenFeatureAcceptanceWindow(IRibbonControl control)
        {
            try
            {
                // Direct feature call keeps this file's global add-in access at one use (the
                // sidebar toggle); the feature opens ownerless, like the seam probe does.
                Features.FeatureAcceptanceWindowFeature.Execute(null);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MainRibbon.OnOpenFeatureAcceptanceWindow");
            }
        }

        public void OnToggleSidebar(IRibbonControl control)
        {
            Globals.ThisAddIn.ToggleSidebar();
        }

        /// <summary>
        /// loadImage callback: returns the custom ribbon icon for controls whose XML sets
        /// image="RibbonIcon". The PNG is embedded (see the csproj EmbeddedResource with the
        /// logical name below) and converted to an IPictureDisp, which the ribbon requires.
        /// </summary>
        public object GetRibbonImage(string imageName)
        {
            try
            {
                if (_ribbonIcon != null) return _ribbonIcon;

                Assembly asm = Assembly.GetExecutingAssembly();
                using (Stream stream = asm.GetManifestResourceStream("oPenEfficiency.Resources.RibbonIcon.png"))
                {
                    if (stream == null) return null;
                    // Keep the Bitmap alive (do not dispose): the OLE picture references it.
                    Bitmap bitmap = new Bitmap(stream);
                    _ribbonIcon = PictureConverter.ImageToPictureDisp(bitmap);
                    return _ribbonIcon;
                }
            }
            catch
            {
                return null;
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }

    /// <summary>
    /// Converts a <see cref="Image"/> to the <c>IPictureDisp</c> that the ribbon loadImage
    /// callback must return. Subclassing <see cref="AxHost"/> exposes its protected
    /// GetIPictureDispFromPicture helper.
    /// </summary>
    internal sealed class PictureConverter : AxHost
    {
        private PictureConverter() : base(string.Empty) { }

        public static object ImageToPictureDisp(Image image) => GetIPictureDispFromPicture(image);
    }
}
