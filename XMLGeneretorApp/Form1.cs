using ExcelParser;
using Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using XMLParser;

namespace XMLGeneretorApp
{
    public partial class Form1 : Form
    {
        public string xmlPath { get; set; }
        public string zipsPath { get; set; }
        public string excelPath { get; set; }
        public string emptyPath { get; set; }
        public string processedExcelsPath { get; set; }
        public string XmlOutput { get; set; }

        public Form1()
        {
            InitializeComponent();
            xmlPath = "C:\\Test\\XMLs\\";
            emptyPath = "C:\\Test\\Empty\\";
            processedExcelsPath = "C:\\Test\\Processed\\";
            XmlOutput = "InsuranceNow_[POLICYNUMBER]_Drivers_[DRIVERS]_Vehicles_[VEHICLES].xml";
            zipsPath = string.Empty;
            excelPath = string.Empty;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            LogForm logForm = new LogForm();
            logForm.ShowDialog();
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
                string zipFullPath = string.Empty;
                string zipName = string.Empty;
                List<Policy> Policies = new List<Policy>();
                List<string> XmlFileNames = new List<string>();
                int total = 1;
                int zippedFiles = 0;
                int zipsCreated = 1;
                String zipsTimestamp = DateTime.Now.ToString("yyyyMMddHHmmssffff");

                ExcelUtil excelUtil = new ExcelUtil(excelPath);
                var workBook = excelUtil.OpenFile();
                excelUtil.ProcessFile(workBook, Policies);
                excelUtil.CloseFile(workBook);

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

                foreach (string xmlFileName in XmlFileNames)
                {

                    if (zippedFiles % 10 == 0)
                    {
                        zipName = "InsuranceNow_" + zipsTimestamp + "_" + zipsCreated.ToString() + ".zip";
                        zipFullPath = zipsPath + zipName;
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

                //string[] splitedPath = excelPath.Split('\\');
                //string excelFileName = splitedPath[splitedPath.Length - 1];
                //string[] newFileData = excelFileName.Split('.');
                //string newFileName = String.Empty;
                //string newFileExtension = newFileData[newFileData.Length - 1];

                //for (int i = 0; i < newFileData.Length - 1; i++)
                //{
                //    newFileName = newFileName + newFileData[i] + ".";
                //}
                //string processedTimeStamp = DateTime.Now.ToString("yyyy-MM-dd h-mm-ss tt");
                //File.Move(excelPath, processedExcelsPath + newFileName + "Processed at " + processedTimeStamp + "." + newFileExtension);
            }
            catch (Exception ex)
            {

            }
        }
    }
}
