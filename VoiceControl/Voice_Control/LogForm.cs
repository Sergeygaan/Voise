using System.Windows.Forms;
using VoiceControlCore;

namespace VoiceControl
{
    public partial class LogForm : Form
    {

        public LogForm()
        {
            InitializeComponent();

            ControlCore controlCore = new ControlCore(listView);
        }
    }
}
