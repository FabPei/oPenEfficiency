using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace oPenEfficiency.UI
{
    public partial class AssetManagerHost : UserControl
    {
        private ElementHost _elementHost;
        private AssetManagerView _wpfView;

        public AssetManagerHost(PowerPointManager manager)
        {
            InitializeComponent(manager ?? throw new System.ArgumentNullException(nameof(manager)));
        }

        private void InitializeComponent(PowerPointManager manager)
        {
            this._elementHost = new ElementHost();
            this._wpfView = new AssetManagerView(manager);
            
            this.SuspendLayout();
            
            this._elementHost.Dock = DockStyle.Fill;
            this._elementHost.Child = this._wpfView;
            
            this.Controls.Add(this._elementHost);
            this.Name = "AssetManagerHost";
            
            this.ResumeLayout(false);
        }

        public void OnSelectionChange(Microsoft.Office.Interop.PowerPoint.Selection selection)
        {
            if (_wpfView != null)
            {
                _wpfView.OnSelectionChange(selection);
            }
        }

        public void SelectSection(string section)
        {
            _wpfView?.SelectSection(section);
        }

        public void SaveCurrentSelection()
        {
            _wpfView?.SaveCurrentSelection();
        }
    }
}
