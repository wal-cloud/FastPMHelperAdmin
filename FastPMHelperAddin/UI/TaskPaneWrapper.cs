using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace FastPMHelperAddin.UI
{
    public class TaskPaneWrapper : UserControl
    {
        private ElementHost _elementHost;

        public TaskPaneWrapper(System.Windows.Controls.UserControl wpfControl)
        {
            _elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = wpfControl
            };

            this.Controls.Add(_elementHost);
        }
    }
}
