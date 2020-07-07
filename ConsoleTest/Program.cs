using ExcelParser;
using Models;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.IO;
//using System.IO.Directory;
using System.Configuration;
using XMLParser;

namespace ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
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
                string ExcelInput = Array.Find(files, f => f.Contains(".xlsm") || f.Contains(".xlsx") || f.Contains(".xls"));
                                               
                string zipFullPath = string.Empty;
                string zipName = string.Empty;
                List<Policy> Policies = new List<Policy>();
                List<string> XmlFileNames = new List<string>();
                int total = 1;
                int zippedFiles = 0;
                int zipsCreated = 1;
                String zipsTimestamp = DateTime.Now.ToString("yyyyMMddHHmmssffff");

                Console.WriteLine("Processing excel file...");
                ExcelUtil excelUtil = new ExcelUtil(ExcelInput);
                var workBook = excelUtil.OpenFile();
                excelUtil.ProcessFile(workBook, Policies);
                excelUtil.CloseFile(workBook);

                Console.WriteLine("Generating XMLs");

                foreach (Policy p in Policies)
                {
                    string fileName = XmlOutput.Replace("[POLICYNUMBER]", p.PolicyNumber.Replace(" ", string.Empty)).
                                                Replace("[DRIVERS]", p.Drivers.Count.ToString()).
                                                Replace("[VEHICLES]", p.Vehicles.Count.ToString());

                    XmlFileNames.Add(fileName);
                    XMLGenerator Generator = new XMLGenerator(xmlPath + fileName, p);
                    Generator.Generate();
                    Console.Write("\r{0}  ", total);
                    total++;
                }
                Console.WriteLine("");
                
                Console.WriteLine("Compressing into zip files");
                
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
                File.Move(ExcelInput, processedExcelsPath + excelFileName);
                
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }

            Console.WriteLine("Done!");
            Console.ReadKey();

        }        
    }
}
