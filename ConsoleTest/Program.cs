using ExcelParser;
using Models;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.IO;
//using System.IO.Directory;
using System.Configuration;
using XMLParser;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace ConsoleTest
{
    class Program : System.Windows.Forms.Form
    {
        private Label label2;
        private TextBox textBox1;
        private TextBox textBox2;
        private RichTextBox richTextBox1;
        private Button button3;
        private Label label3;
        private Button button4;
        private OpenFileDialog openFileDialog1;
        private Label label1;

        
        private Button button1;
        private Button button2;
        
        
        private string processedExcelsPath;
        private string xmlPath;
    
        Program() // ADD THIS CONSTRUCTOR
        {
            InitializeComponent();
        }

        static void Main()
        {
            try
            {
                string xmlPath = Convert.ToString(ConfigurationManager.AppSettings.Get("xmlPath"));
                string zipsPath = Convert.ToString(ConfigurationManager.AppSettings.Get("zipsPath"));
                string excelPath = Convert.ToString(ConfigurationManager.AppSettings.Get("excelPath"));
                string emptyPath = Convert.ToString(ConfigurationManager.AppSettings.Get("emptyPath"));
                string processedExcelsPath = Convert.ToString(ConfigurationManager.AppSettings.Get("processedExcelsPath"));

                string XmlOutput = "InsuranceNow_[POLICYNUMBER]_Drivers_[DRIVERS]_Vehicles_[VEHICLES].xml";

                string[] files = Directory.GetFiles(excelPath, "*", SearchOption.AllDirectories);
                string ExcelInput = Array.Find(files, f => f.Contains(excelPath) && !f.Contains(processedExcelsPath) && (f.Contains(".xlsm") || f.Contains(".xlsx") || f.Contains(".xls")));
                                               
                string zipFullPath = string.Empty;
                string zipName = string.Empty;
                List<Policy> Policies = new List<Policy>();
                List<string> XmlFileNames = new List<string>();
                int total = 1;
                int zippedFiles = 0;
                int zipsCreated = 1;
                String zipsTimestamp = DateTime.Now.ToString("yyyyMMddHHmmssffff");

                Application.EnableVisualStyles();

                Application.Run(new Program());

                //Console.WriteLine("Processing excel file...");
                ExcelUtil excelUtil = new ExcelUtil(ExcelInput);
                var workBook = excelUtil.OpenFile();
                excelUtil.ProcessFile(workBook, Policies);
                excelUtil.CloseFile(workBook);

                //Console.WriteLine("Generating XMLs");

                foreach (Policy p in Policies)
                {
                    string fileName = XmlOutput.Replace("[POLICYNUMBER]", p.PolicyNumber.Replace(" ", string.Empty)).
                                                Replace("[DRIVERS]", p.Drivers.Count.ToString()).
                                                Replace("[VEHICLES]", p.Vehicles.Count.ToString());

                    XmlFileNames.Add(fileName);
                    XMLGenerator Generator = new XMLGenerator(xmlPath + fileName, p);
                    Generator.Generate();
                    //Console.Write("\r{0}  ", total);
                    total++;
                }
                //Console.WriteLine("");
                
                //Console.WriteLine("Compressing into zip files");
                
                foreach(string xmlFileName in XmlFileNames)
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

                string[] splitedPath = ExcelInput.Split('\\');
                string excelFileName = splitedPath[splitedPath.Length - 1];
                string[] newFileData = excelFileName.Split('.');
                string newFileName = String.Empty;
                string newFileExtension = newFileData[newFileData.Length - 1];

                for (int i = 0; i < newFileData.Length - 1; i++)
                {
                    newFileName = newFileName + newFileData[i] + ".";
                }
                string processedTimeStamp = DateTime.Now.ToString("yyyy-MM-dd h-mm-ss tt");
                File.Move(ExcelInput, processedExcelsPath + newFileName + "Processed at " + processedTimeStamp + "." + newFileExtension);
                
            }
            catch (Exception e)
            {
                //Console.WriteLine(e.Message);
                //Console.WriteLine(e.StackTrace);
            }

            //Console.WriteLine("Done!");
            //Console.ReadKey();

        }

        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(82, 95);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "Excel File";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(30, 137);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(119, 17);
            this.label2.TabIndex = 1;
            this.label2.Text = "Destination folder";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // textBox1
            // 
            this.textBox1.Enabled = false;
            this.textBox1.Location = new System.Drawing.Point(155, 90);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(339, 22);
            this.textBox1.TabIndex = 4;
            // 
            // textBox2
            // 
            this.textBox2.Enabled = false;
            this.textBox2.Location = new System.Drawing.Point(155, 131);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(339, 22);
            this.textBox2.TabIndex = 5;
            // 
            // richTextBox1
            // 
            this.richTextBox1.Enabled = false;
            this.richTextBox1.Location = new System.Drawing.Point(33, 253);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(542, 102);
            this.richTextBox1.TabIndex = 6;
            this.richTextBox1.Text = "";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(213, 180);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(230, 38);
            this.button3.TabIndex = 7;
            this.button3.Text = "Convert to XML and Compress";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(30, 233);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(52, 17);
            this.label3.TabIndex = 8;
            this.label3.Text = "Status ";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(500, 361);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 26);
            this.button4.TabIndex = 9;
            this.button4.Text = "Log";
            this.button4.UseVisualStyleBackColor = true;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(501, 88);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 10;
            this.button1.Text = "Browse";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(501, 130);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 11;
            this.button2.Text = "Browse";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Program
            // 
            this.ClientSize = new System.Drawing.Size(611, 409);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "Program";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }        

        private void button3_Click(object sender, EventArgs e)
        {
                             
        }

        private void OpenExcelPathFileDialog()
        {
            openFileDialog1.ShowDialog();
            string fileName = openFileDialog1.FileName;
            string readFile = File.ReadAllText(fileName);
            this.processedExcelsPath = fileName;
            SetStatusText(fileName);
        }

        private void OpendDestPathFileDialog()
        {
            openFileDialog1.ShowDialog();
            string fileName = openFileDialog1.FileName;
            string readFile = File.ReadAllText(fileName);
            this.xmlPath = fileName;
            SetStatusText(fileName);
        }

        private void SetStatusText(string text)
        {
            richTextBox1.Invoke((Action)delegate
            {
                richTextBox1.Text = text;
            });
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Thread newThread = new Thread(new ThreadStart(OpenExcelPathFileDialog));
            newThread.SetApartmentState(ApartmentState.STA);
            newThread.Start();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Thread newThread = new Thread(new ThreadStart(OpendDestPathFileDialog));
            newThread.SetApartmentState(ApartmentState.STA);
            newThread.Start();
        }        
    }
}
