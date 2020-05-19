using InsuranceNow_XMLGenerator.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsuranceNow_XMLGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string path = "C:\\Test\\";
                string XmlOutput = "InsuranceNow_[POLICYNUMBER].xml";
                string ExcelInput = "_dataMigrationVer1.5.xlsx";
                //string ExcelInput = path + "Test.xlsm";
                List<Policy> Policies = new List<Policy>();
                int total = 1;

                Console.WriteLine("Processing excel file...");
                ExcelUtil excelUtil = new ExcelUtil(path + ExcelInput);
                var workBook = excelUtil.OpenFile();
                excelUtil.ProcessFile(workBook, Policies);
                excelUtil.CloseFile(workBook);

                Console.WriteLine("Generating XMLs");

                foreach(Policy p in Policies)
                {
                    string fileName = XmlOutput.Replace("[POLICYNUMBER]", p.PolicyNumber.Replace(" ", string.Empty));
                    XMLGenerator Generator = new XMLGenerator(path + "XMLs\\" + fileName, p);
                    Generator.Generate();
                    Console.Write("\r{0}  ", total);
                    total++;
                }
                Console.WriteLine("");
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

            Console.WriteLine("Done!");
            Console.ReadKey();
            
        }
    }
}
