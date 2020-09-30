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
    public partial class PublicKey : Form
    {
        public PublicKey()
        {
            InitializeComponent();
            rtbPublickKey.Text = Properties.Settings.Default["PublicKey"] != null ? Properties.Settings.Default["PublicKey"].ToString() : String.Empty;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (String.IsNullOrEmpty(rtbPublickKey.Text))
                {
                    MessageBox.Show("The Public Key cannot be empty.");
                }

                Properties.Settings.Default["PublicKey"] = rtbPublickKey.Text;
                Properties.Settings.Default.Save();

                MessageBox.Show("The Public Key has been saved.");
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was an error saving the public key: " + ex.Message);
            }
        }
    }
}
