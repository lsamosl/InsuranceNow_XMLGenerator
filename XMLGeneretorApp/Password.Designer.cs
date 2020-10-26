namespace XMLGeneretorApp
{
    partial class Password
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tbPassword = new System.Windows.Forms.TextBox();
            this.tbConfirmPassword = new System.Windows.Forms.TextBox();
            this.buttonSetPassword = new System.Windows.Forms.Button();
            this.alert = new System.Windows.Forms.Label();
            this.enablePassCheck = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(89, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(69, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "Password";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(38, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(120, 17);
            this.label2.TabIndex = 1;
            this.label2.Text = "Confirm password";
            // 
            // tbPassword
            // 
            this.tbPassword.Location = new System.Drawing.Point(164, 41);
            this.tbPassword.Name = "tbPassword";
            this.tbPassword.PasswordChar = '*';
            this.tbPassword.Size = new System.Drawing.Size(207, 22);
            this.tbPassword.TabIndex = 2;
            // 
            // tbConfirmPassword
            // 
            this.tbConfirmPassword.Location = new System.Drawing.Point(164, 76);
            this.tbConfirmPassword.Name = "tbConfirmPassword";
            this.tbConfirmPassword.PasswordChar = '*';
            this.tbConfirmPassword.Size = new System.Drawing.Size(207, 22);
            this.tbConfirmPassword.TabIndex = 3;
            // 
            // buttonSetPassword
            // 
            this.buttonSetPassword.Location = new System.Drawing.Point(135, 127);
            this.buttonSetPassword.Name = "buttonSetPassword";
            this.buttonSetPassword.Size = new System.Drawing.Size(182, 24);
            this.buttonSetPassword.TabIndex = 4;
            this.buttonSetPassword.Text = "Set and save password";
            this.buttonSetPassword.UseVisualStyleBackColor = true;
            this.buttonSetPassword.Click += new System.EventHandler(this.buttonSetPassword_Click);
            // 
            // alert
            // 
            this.alert.AutoSize = true;
            this.alert.ForeColor = System.Drawing.Color.Red;
            this.alert.Location = new System.Drawing.Point(54, 101);
            this.alert.Name = "alert";
            this.alert.Size = new System.Drawing.Size(324, 17);
            this.alert.TabIndex = 5;
            this.alert.Text = "*Password and *Confirm password does not match";
            this.alert.Visible = false;
            this.alert.Click += new System.EventHandler(this.alert_Click);
            // 
            // enablePassCheck
            // 
            this.enablePassCheck.AutoSize = true;
            this.enablePassCheck.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.enablePassCheck.Checked = true;
            this.enablePassCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.enablePassCheck.Location = new System.Drawing.Point(41, 12);
            this.enablePassCheck.Name = "enablePassCheck";
            this.enablePassCheck.Size = new System.Drawing.Size(138, 21);
            this.enablePassCheck.TabIndex = 6;
            this.enablePassCheck.Text = "Enable password";
            this.enablePassCheck.UseVisualStyleBackColor = true;
            this.enablePassCheck.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // Password
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(444, 163);
            this.Controls.Add(this.enablePassCheck);
            this.Controls.Add(this.alert);
            this.Controls.Add(this.buttonSetPassword);
            this.Controls.Add(this.tbConfirmPassword);
            this.Controls.Add(this.tbPassword);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "Password";
            this.Text = "Password";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbPassword;
        private System.Windows.Forms.TextBox tbConfirmPassword;
        private System.Windows.Forms.Button buttonSetPassword;
        private System.Windows.Forms.Label alert;
        private System.Windows.Forms.CheckBox enablePassCheck;
    }
}