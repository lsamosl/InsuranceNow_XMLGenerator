using ExcelParser;
using Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading.Tasks;
using XMLParser;

namespace XMLGeneretorApp
{
    public partial class frmGenerate : Form
    {
        public string xmlPath { get; set; }
        public string zipsPath { get; set; }
        public string excelPath { get; set; }
        public string emptyPath { get; set; }
        public string processedExcelsPath { get; set; }
        public string XmlOutput { get; set; }

        private string version = Convert.ToString(ConfigurationManager.AppSettings.Get("version"));


        public frmGenerate()
        {
            InitializeComponent();
            xmlPath = "C:\\Test\\XMLs\\";
            emptyPath = "C:\\Test\\Empty\\";
            processedExcelsPath = "C:\\Test\\Processed\\";
            XmlOutput = "InsuranceNow_[POLICYNUMBER]_Drivers_[DRIVERS]_Vehicles_[VEHICLES].xml";
            zipsPath = string.Empty;
            excelPath = string.Empty;
            versionLabel.Text = version;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //LogForm logForm = new LogForm();
            //logForm.ShowDialog();
            System.Diagnostics.Process.Start("notepad.exe", "C:\\Test\\SafeNetLog.txt");
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
            }                
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                statusTb.Text = string.Empty;
                button2.Enabled = false;
                button3.Enabled = false;
                button4.Enabled = false;

                backgroundWorker1.RunWorkerAsync();
            }
            catch (Exception ex)
            {
                errorLabel.Visible = true;
            }
        }

        private void Generate()
        {
            Action action = null;
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
                    XMLGenerator Generator = new XMLGenerator(xmlPath + fileName, p);
                    Generator.Generate();
                    total++;
                }

                var recordsReaded = total.ToString() + " policies readed";

                action = () => statusTb.AppendText(Environment.NewLine + "Compressing into zip files...");
                statusTb.Invoke(action);                            

                foreach (string xmlFileName in XmlFileNames)
                {

                    if (zippedFiles % 10 == 0)
                    {
                        zipName = "InsuranceNow_" + zipsTimestamp + "_" + zipsCreated.ToString() + ".zip";
                        zipFullPath = zipsPath + "\\" + zipName;
                        ZipFile.CreateFromDirectory(emptyPath, zipFullPath);
                        zipsCreated++;
                    }

                    using (ZipArchive zip = ZipFile.Open(zipFullPath, ZipArchiveMode.Update))
                    {
                        string filePath = xmlPath + xmlFileName;
                        zip.CreateEntryFromFile(filePath, xmlFileName);
                        zippedFiles++;
                        File.Delete(filePath);
                    }
                }

                action = () => statusTb.AppendText(Environment.NewLine + "Process completed successfully at "+DateTime.Now);
                statusTb.Invoke(action);                
            }
            catch(Exception ex)
            {
                action = () => statusTb.AppendText(Environment.NewLine + "An error occurred: " + ex.Message);
                statusTb.Invoke(action);

                action = () => statusTb.AppendText(Environment.NewLine + "The process ended with errors at " + DateTime.Now);
                statusTb.Invoke(action);
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
        }
    }
}
