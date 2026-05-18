using System.Windows.Forms.Integration;
using System.Windows.Forms;

namespace oPenEfficiency
{
    /// <summary>
    /// WinForms-UserControl, das das WPF-MainSidebar in einem ElementHost für das VSTO Task Pane hostet.
    /// </summary>
    public class SidebarTaskPaneHost : UserControl
    {
        private readonly ElementHost _elementHost;

        public SidebarTaskPaneHost()
        {
            _elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = new MainSidebar()
            };
            Controls.Add(_elementHost);
        }

        /// <summary>
        /// Zugriff auf das WPF-Sidebar-Control (z. B. für Event-Handler oder Datenbindung).
        /// </summary>
        public MainSidebar Sidebar => _elementHost.Child as MainSidebar;
    }
}
