using ExcelParser;
using Models;
using System;
using System.Collections.Generic;
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
                string ExcelInput = "_dataMigrationVer1.5.xlsm";
                List<Policy> Policies = new List<Policy>();
                int total = 1;

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

                    XMLGenerator Generator = new XMLGenerator(path + "XMLs\\" + fileName, p);
                    Generator.Generate();
                    Console.Write("\r{0}  ", total);
                    total++;
                }
                Console.WriteLine("");
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
