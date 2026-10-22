using oPenEfficiency.UI.Dialogs;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI.Panels
{
    internal interface ISidebarHost
    {
        PowerPoint.Application Application { get; }
        bool IsAssetManagerVisible { get; }
        void OpenAssetManager(string section = null, bool saveCurrentSelection = false);
        void ToggleAssetManager();
        void OpenFeatureExplorer(string initialCategory = null, bool focusSearch = false);
        void OpenIconBrowser(IconBrowserStartView startView = IconBrowserStartView.Default);
        void ToggleThemeColorPane();
        void ConfigurePerformanceMode();
    }
}
