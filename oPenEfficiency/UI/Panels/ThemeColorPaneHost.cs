using System.Windows.Forms.Integration;
using System.Windows.Forms;
using oPenEfficiency.UI;

namespace oPenEfficiency
{
    public class ThemeColorPaneHost : UserControl
    {
        private readonly ElementHost _elementHost;
        private readonly ThemeColorControl _control;

        public ThemeColorPaneHost(PowerPointManager manager)
        {
            _control = new ThemeColorControl();
            _control.Initialize(manager, true);

            _elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = _control
            };
            Controls.Add(_elementHost);
        }
    }
}
