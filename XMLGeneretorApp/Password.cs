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
    public partial class Password : Form
    {

        public string password { get; set; }        

        public Password()
        {
            InitializeComponent();
            var zipsPasswordEnabled = Convert.ToBoolean(Properties.Settings.Default["zipsPasswordEnabled"].ToString());

            enablePassCheck.Checked = zipsPasswordEnabled;
                       
            if (zipsPasswordEnabled)
            {
                tbPassword.Text = Properties.Settings.Default["zipsPassword"] != null ? Properties.Settings.Default["zipsPassword"].ToString() : String.Empty;
                tbConfirmPassword.Text = Properties.Settings.Default["zipsPassword"] != null ? Properties.Settings.Default["zipsPassword"].ToString() : String.Empty;
            }
            else
            {
                tbPassword.Enabled = false;
                tbConfirmPassword.Enabled = false;
                buttonSetPassword.Enabled = false;
            }

            
            password = tbPassword.Text;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void alert_Click(object sender, EventArgs e)
        {

        }

        private void buttonSetPassword_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tbPassword.Text) && string.IsNullOrEmpty(tbConfirmPassword.Text))
            {
                MessageBox.Show("Password cannot be blank");
            }
            else
            {
                if (tbPassword.Text.Equals(tbConfirmPassword.Text))
                {
                    password = tbPassword.Text;

                    if (alert.Visible)
                    {
                        alert.Visible = false;
                    }

                    Properties.Settings.Default["zipsPassword"] = tbPassword.Text;
                    Properties.Settings.Default.Save();

                    MessageBox.Show("Password has been saved");
                    this.Close();
                }
                else
                {
                    alert.Visible = true;
                }
            }

                        
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["zipsPasswordEnabled"] = enablePassCheck.Checked;

            if(enablePassCheck.Checked)
            {                 
                tbPassword.Enabled = true;
                tbConfirmPassword.Enabled = true;
                buttonSetPassword.Enabled = true;
            }
            else
            {
                tbPassword.Enabled = false;
                tbConfirmPassword.Enabled = false;
                buttonSetPassword.Enabled = false;
            }
            
        }
    }
}
