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
                string XmlOutput = path + "InsuranceNow_" + Guid.NewGuid().ToString() + ".xml";
                string ExcelInput = path + "_dataMigrationVer1.5.xlsx";
                //string ExcelInput = path + "Test.xlsm";
                List<Policy> Policies = new List<Policy>();

                Console.WriteLine("Processing excel file...");
                ExcelUtil excelUtil = new ExcelUtil(ExcelInput);
                var workBook = excelUtil.OpenFile();
                excelUtil.ProcessFile(workBook, Policies);
                excelUtil.CloseFile(workBook);

                Console.WriteLine("Generating XML");
                XMLGenerator Generator = new XMLGenerator(XmlOutput, Policies);
                Generator.Generate();
                
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
