using ExcelParser;
using Models;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.IO;
//using System.IO.Compression.FileSystem;
using XMLParser;

namespace ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string path = "C:\\Test\\";
                string XmlOutput = "InsuranceNow_[POLICYNUMBER]_Drivers_[DRIVERS]_Vehicles_[VEHICLES].xml";
                string ExcelInput = "dataMigrationVer1.6.xlsm";
                string zipPath = string.Empty;
                string zipFullPath = string.Empty;
                List<Policy> Policies = new List<Policy>();
                List<string> XmlFileNames = new List<string>();
                int total = 1;
                int zippedFiles = 0;
                int zipsCreated = 1;

                Console.WriteLine("Processing excel file...");
                ExcelUtil excelUtil = new ExcelUtil(path + ExcelInput);
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
                    XMLGenerator Generator = new XMLGenerator(path + "XMLs\\" + fileName, p);
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
                        string startPath = path + "Empty\\";
                        zipPath = "InsuranceNow_" + zipsCreated.ToString() + ".zip";
                        zipFullPath = "C:\\Test\\Zips\\" + zipPath;
                        //File.Delete(zipFullPath);
                        ZipFile.CreateFromDirectory(startPath, zipFullPath);                        
                        zipsCreated++;
                    }

                    using (ZipArchive zip = ZipFile.Open(zipFullPath, ZipArchiveMode.Update))
                    {
                        string filePath = "C:\\Test\\XMLs\\" + xmlFileName;
                        zip.CreateEntryFromFile(filePath, xmlFileName);
                        zippedFiles++;
                    }                    
                }
                
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
