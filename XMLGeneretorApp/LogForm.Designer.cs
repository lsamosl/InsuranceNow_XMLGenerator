namespace XMLGeneretorApp
{
    partial class LogForm
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
            this.logTextArea = new System.Windows.Forms.RichTextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // logTextArea
            // 
            this.logTextArea.AccessibleName = "taLog";
            this.logTextArea.Enabled = false;
            this.logTextArea.Location = new System.Drawing.Point(12, 39);
            this.logTextArea.Name = "logTextArea";
            this.logTextArea.Size = new System.Drawing.Size(528, 232);
            this.logTextArea.TabIndex = 0;
            this.logTextArea.Text = "";
            this.logTextArea.TextChanged += new System.EventHandler(this.logTextArea_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "Log entries";
            // 
            // LogForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(552, 295);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.logTextArea);
            this.Name = "LogForm";
            this.Text = "LogForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox logTextArea;
        private System.Windows.Forms.Label label1;
    }
}