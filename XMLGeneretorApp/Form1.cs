using ExcelParser;
using Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
//using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using XMLParser;
using Cinchoo.PGP;
using Ionic.Zip;


namespace XMLGeneretorApp
{
    public partial class frmGenerate : Form
    {
        public string xmlPath { get; set; }
        public string zipsPath { get; set; }
        public string PGPPath { get; set; }
        public string excelPath { get; set; }
        public string emptyPath { get; set; }
        public string processedExcelsPath { get; set; }
        public string XmlOutput { get; set; }
        public ChoPGPEncryptDecrypt PGP { get; set; }
        public string PublicKey { get; set; }
        public FileInfo PublicKeyPath { get; set; }
        public string zipsPassword { get; set; }
        public bool enableZipsPassword { get; set; }

        //private string version = Convert.ToString(ConfigurationManager.AppSettings.Get("version"));
        private string version = System.Windows.Forms.Application.ProductVersion;

        public frmGenerate()
        {
            InitializeComponent();
            xmlPath = "XMLs\\";
            emptyPath = "Empty\\";
            processedExcelsPath = "Processed\\";
            XmlOutput = "InsuranceNow_[POLICYNUMBER]_Drivers_[DRIVERS]_Vehicles_[VEHICLES].xml";
            zipsPath = string.Empty;
            excelPath = string.Empty;
            PGPPath = string.Empty;
            versionLabel.Text = version;
            PGP = new ChoPGPEncryptDecrypt();
            PublicKey = Properties.Settings.Default["PublicKey"].ToString();
            zipsPassword = Properties.Settings.Default["zipsPassword"].ToString();
            enableZipsPassword = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //LogForm logForm = new LogForm();
            //logForm.ShowDialog();
            var path = ConfigurationManager.AppSettings.Get("LogFileName");
            if (!File.Exists(path))
            {
                File.WriteAllText(path, String.Empty);
            }
            System.Diagnostics.Process.Start("notepad.exe", path);
            return;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            excelBrowse.ShowDialog();
        }

        private void excelBrowse_FileOk(object sender, CancelEventArgs e)
        {
            tbExcelPath.Text = excelBrowse.FileName;
            excelPath = tbExcelPath.Text;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var result = outputBrowse.ShowDialog();
            if (result == DialogResult.OK)
            {
                tbOutput.Text = outputBrowse.SelectedPath;
                zipsPath = outputBrowse.SelectedPath;
                //passwordToolStripMenuItem.Enabled = true;
            }                        
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                PublicKey = Properties.Settings.Default["PublicKey"].ToString();
                zipsPassword = Properties.Settings.Default["zipsPassword"].ToString();

                if (string.IsNullOrEmpty(PublicKey))
                {
                    MessageBox.Show("The public key has not been saved.");
                    return;
                }
                else
                {
                    string filename = Path.GetTempFileName();

                    PublicKeyPath = new FileInfo(filename);
                    PublicKeyPath.Attributes = FileAttributes.Temporary;

                    using (StreamWriter sw = new StreamWriter(filename))
                        sw.WriteLine(PublicKey);
                }

                if (enableZipsPassword)
                {
                    if (string.IsNullOrEmpty(zipsPassword))
                    {
                        MessageBox.Show("A password is needed");
                        return;
                    }
                }
                
                statusTb.Text = string.Empty;
                button2.Enabled = false;
                button3.Enabled = false;
                button4.Enabled = false;
                btnPGPBrowse.Enabled = false;

                backgroundWorker1.RunWorkerAsync();
            }
            catch
            {
                errorLabel.Visible = true;
            }
        }

        private void Generate()
        {
            Action action = null;
            Dictionary<String, String> ZipFiles = new Dictionary<String, String>();

            try
            {
                string zipFullPath = string.Empty;
                string zipName = string.Empty;
                List<Policy> Policies = new List<Policy>();
                List<string> XmlFileNames = new List<string>();
                int total = 1;
                int zippedFiles = 0;
                int zipsCreated = 1;
                String zipsTimestamp = DateTime.Now.ToString("yyyyMMddHHmmssffff");                

                action = () => passwordToolStripMenuItem.Enabled = false;
                statusTb.Invoke(action);

                action = () => enablePasswordZipsToolStripMenuItem.Enabled = false;
                statusTb.Invoke(action);

                action = () => statusTb.AppendText("Process started at "+DateTime.Now);
                statusTb.Invoke(action);                

                action = () => statusTb.AppendText(Environment.NewLine + "Reading Excel file...");
                statusTb.Invoke(action);

                ExcelUtil excelUtil = new ExcelUtil(excelPath);
                var workBook = excelUtil.OpenFile();
                excelUtil.ProcessFile(workBook, Policies);
                excelUtil.CloseFile(workBook);

                action = () => statusTb.AppendText(Environment.NewLine + "Generating XMLs...");
                statusTb.Invoke(action); 

                foreach (Policy p in Policies)
                {
                    string fileName = XmlOutput.Replace("[POLICYNUMBER]", p.PolicyNumber.Replace(" ", string.Empty)).
                                                Replace("[DRIVERS]", p.Drivers.Count.ToString()).
                                                Replace("[VEHICLES]", p.Vehicles.Count.ToString());

                    XmlFileNames.Add(fileName);
                    if (!Directory.Exists(xmlPath))
                    {
                        Directory.CreateDirectory(xmlPath);
                    }
                    XMLGenerator Generator = new XMLGenerator(xmlPath + fileName, p);
                    Generator.Generate();
                    total++;
                }

                var recordsReaded = total.ToString() + " policies readed";

                action = () => statusTb.AppendText(Environment.NewLine + "Compressing into zip files...");
                statusTb.Invoke(action);

                if (!Directory.Exists(emptyPath))
                {
                    Directory.CreateDirectory(emptyPath);
                }

                foreach (string xmlFileName in XmlFileNames)
                {

                    if (zippedFiles % 10 == 0)
                    {
                        zipName = "InsuranceNow_" + zipsTimestamp + "_" + zipsCreated.ToString() + ".zip";
                        zipFullPath = zipsPath + "\\" + zipName;
                        ZipFiles.Add(zipName, zipFullPath);

                        using (ZipFile zip = new ZipFile(zipFullPath))
                        {
                            if (enableZipsPassword)
                            {
                                zip.Password = zipsPassword;
                            }                            
                            zip.Save();
                        }

                        zipsCreated++;
                    }

                    using (ZipFile zip = ZipFile.Read(zipFullPath))
                    {
                        string filePath = xmlPath + xmlFileName;

                        if (enableZipsPassword)
                        {
                            zip.Password = zipsPassword;
                        }
                                                
                        zip.AddFile(filePath, String.Empty);
                        zip.Save();
                        zippedFiles++;

                        File.Delete(filePath);
                    }
                }

                action = () => statusTb.AppendText(Environment.NewLine + "Encrypting Zip Files... ");
                statusTb.Invoke(action);

                if (!Directory.Exists(PGPPath))
                    Directory.CreateDirectory(PGPPath);

                foreach (var zipfile in ZipFiles)
                {
                    string outputPgpFile = string.Format("{0}\\{1}.pgp", PGPPath, zipfile.Key);
                    PGP.EncryptFile(zipfile.Value, outputPgpFile, PublicKeyPath.FullName, false, true);
                }
                    
                action = () => statusTb.AppendText(Environment.NewLine + "Process completed successfully at "+DateTime.Now);
                statusTb.Invoke(action);

                action = () => enablePasswordZipsToolStripMenuItem.Enabled = true;
                statusTb.Invoke(action);

                action = () => passwordToolStripMenuItem.Enabled = true;
                statusTb.Invoke(action);
            }
            catch(Exception ex)
            {
                action = () => statusTb.AppendText(Environment.NewLine + "An error occurred: " + ex.Message);
                statusTb.Invoke(action);

                action = () => statusTb.AppendText(Environment.NewLine + "The process ended with errors at " + DateTime.Now);
                statusTb.Invoke(action);

                var path = ConfigurationManager.AppSettings.Get("LogFileName");
                if (!File.Exists(path))
                {
                    File.WriteAllText(path, "An Error ocurred: " + ex.Message + " at " + DateTime.Now.ToString() + Environment.NewLine);
                }

                String previousText = String.Empty;
                previousText = File.ReadAllText(path);

                File.WriteAllText(path, String.Empty);
                File.AppendAllText(path, "An Error ocurred: " + ex.Message + " at " + DateTime.Now.ToString() + Environment.NewLine);
                File.AppendAllText(path, previousText);                                
            }
            finally
            {
                if (!backgroundWorker1.IsBusy)
                    backgroundWorker1.CancelAsync();
            }
        }

        private void statusTb_TextChanged(object sender, EventArgs e)
        {

        }

        private void errorLabel_Click(object sender, EventArgs e)
        {

        }

        private void UpdateStatusTB(string newLine)
        {
            statusTb.Text += Environment.NewLine;
            statusTb.Text += newLine;
        }

        private void UpdateGenerationButton()
        {
            if (tbExcelPath.Text != String.Empty && tbOutput.Text != String.Empty)
            {
                button4.Enabled = true;
            }
            else
            {
                button4.Enabled = false;
            }
        }

        private void tbExcelPath_TextChanged(object sender, EventArgs e)
        {
            UpdateGenerationButton();
        }

        private void tbOutput_TextChanged(object sender, EventArgs e)
        {
            PGPPath = tbOutput.Text + "\\PGP_Files";
            tbOutputPGP.Text = PGPPath;
            UpdateGenerationButton();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker helperBW = sender as BackgroundWorker;
            Generate();

            if (helperBW.CancellationPending)
            {
                e.Cancel = true;
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            btnPGPBrowse.Enabled = true;
        }

        private void btnPGPBrowse_Click(object sender, EventArgs e)
        {
            var result = outputPGPBrowse.ShowDialog();
            if (result == DialogResult.OK)
            {
                tbOutputPGP.Text = outputPGPBrowse.SelectedPath;
                PGPPath = outputPGPBrowse.SelectedPath;
            }
        }

        private void publicKeyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PublicKey form = new PublicKey();
            form.ShowDialog();
            return;
        }        

        private void passwordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Password form = new Password();
            form.ShowDialog();
            return;
        }

        private void enablePasswordZipsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            enableZipsPassword = enablePasswordZipsToolStripMenuItem.Checked;
            passwordToolStripMenuItem.Enabled = enableZipsPassword;            
        }        
    }
}
