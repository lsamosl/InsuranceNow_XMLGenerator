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
                string fileName = "InsuranceNow_" + Guid.NewGuid().ToString() + ".xml";
                string path = "C:\\Test\\" + fileName;
                XMlGenerator Generator = new XMlGenerator(path);
                Generator.Generate();
                Console.WriteLine("Done!");
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

            Console.ReadKey();
            
        }
    }
}
