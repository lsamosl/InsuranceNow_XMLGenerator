using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XMLGeneretorApp
{
    public partial class LogForm : Form
    {
        public LogForm()
        {
            InitializeComponent();
        }

        public static string taLog
        {
            get { return taLog; }
            set { taLog = value; }
        }

        private void logTextArea_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
