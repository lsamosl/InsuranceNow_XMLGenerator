using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using InsuranceNow_XMLGenerator.Models;
using InsuranceNow_XMLGenerator.Resources;
using Microsoft.Office.Interop.Excel;

namespace InsuranceNow_XMLGenerator
{
    public class ExcelUtil
    {
        int DRIVER_FIRSTNAME, DRIVER_MIDDLENAME, DRIVER_LASTNAME, DRIVER_DRIVERNUMBER, 
            DRIVER_POLICY, VEHICLE_MAKE, VEHICLE_MODEL, VEHICLE_VIN, VEHICLE_POLICY;

        public string FileName { get; set; }
        public int IndexRow { get; set; }

        public ExcelUtil(string fileName)
        {
            FileName = fileName;
            IndexRow = 2;
            DRIVER_FIRSTNAME = GetExcelColumnNumber(ExcelResources.Driver_FirstName);
            DRIVER_MIDDLENAME = GetExcelColumnNumber(ExcelResources.Driver_MiddleName);
            DRIVER_LASTNAME = GetExcelColumnNumber(ExcelResources.Driver_LastName);
            DRIVER_DRIVERNUMBER = GetExcelColumnNumber(ExcelResources.Driver_DriverNumber);
            DRIVER_POLICY = GetExcelColumnNumber(ExcelResources.Driver_Policy);
            VEHICLE_MAKE = GetExcelColumnNumber(ExcelResources.Vehicle_Make);
            VEHICLE_MODEL = GetExcelColumnNumber(ExcelResources.Vehicle_Model);
            VEHICLE_VIN = GetExcelColumnNumber(ExcelResources.Vehicle_VIN);
            VEHICLE_POLICY = GetExcelColumnNumber(ExcelResources.Vehicle_Policy);
        }

        public Workbook OpenFile()
        {
            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(FileName, ReadOnly: false);

            return xlWorkbook;
        }

        public void ProcessFile(Workbook xlWorkbook, List<Policy> Policies)
        {
            try
            {
                _Worksheet PolicySheet = xlWorkbook.Sheets[ExcelResources.PolicySheet];
                _Worksheet DriverSheet = xlWorkbook.Sheets[ExcelResources.DriverSheet];
                _Worksheet VehicleSheet = xlWorkbook.Sheets[ExcelResources.VehicleSheet];

                Range rangeObject = PolicySheet.Cells[IndexRow, ExcelResources.Policy_PolicyNumber];
                string policyNumber = (string)rangeObject.Value2;

                Policy policy = new Policy();

                do
                {
                    policy = new Policy {
                        PolicyNumber = policyNumber,
                        Drivers = new List<Driver>(),
                        Vehicles = new List<Vehicle>()
                    };

                    Range xlrangeDriver = DriverSheet.UsedRange;
                    xlrangeDriver.AutoFilter(GetExcelColumnNumber(ExcelResources.Driver_Policy), policyNumber, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                    Range filteredRangeDriver = xlrangeDriver.SpecialCells(XlCellType.xlCellTypeVisible, XlSpecialCellsValue.xlTextValues);

                    for (int areaId = 1; areaId <= filteredRangeDriver.Areas.Count; areaId++)
                    {
                        Range areaRange = filteredRangeDriver.Areas.get_Item(areaId);
                        object[,] areaValues = areaRange.Value2;

                        for (int i = 2; i <= areaValues.GetLength(0) ; i++)
                        {
                            policy.Drivers.Add(new Driver
                            {
                                DriverNumber = Int32.Parse(areaValues[i, DRIVER_DRIVERNUMBER].ToString()),
                                FirstName = areaValues[i, DRIVER_FIRSTNAME].ToString(),
                                MiddleName = areaValues[i, DRIVER_MIDDLENAME].ToString(),
                                LastName = areaValues[i, DRIVER_LASTNAME].ToString(),
                            }) ;
                        }
                    }

                    Range xlrangeVehicle = VehicleSheet.UsedRange;
                    xlrangeVehicle.AutoFilter(GetExcelColumnNumber(ExcelResources.Vehicle_Policy), policyNumber, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                    Range filteredRangeVehicle = xlrangeVehicle.SpecialCells(XlCellType.xlCellTypeVisible, XlSpecialCellsValue.xlTextValues);

                    for (int areaId = 1; areaId <= filteredRangeVehicle.Areas.Count; areaId++)
                    {
                        Range areaRange = filteredRangeVehicle.Areas.get_Item(areaId);
                        object[,] areaValues = areaRange.Value2;

                        for (int i = 2; i <= areaValues.GetLength(0); i++)
                        {
                            policy.Vehicles.Add(new Vehicle
                            {
                                Make = areaValues[i, VEHICLE_MAKE].ToString(),
                                Model = areaValues[i, VEHICLE_MODEL].ToString(),
                                VIN = areaValues[i, VEHICLE_VIN].ToString()
                            });
                        }
                    }

                    Policies.Add(policy);

                    IndexRow++;
                    rangeObject = PolicySheet.Cells[IndexRow, ExcelResources.Policy_PolicyNumber];
                    policyNumber = (string)rangeObject.Value2;
                }
                while (!string.IsNullOrEmpty(policyNumber));
            }
            catch(Exception e)
            {
                CloseFile(xlWorkbook);
                throw e;
            }
        }

        public void CloseFile(Workbook xlWorkbook)
        {
            xlWorkbook.Close(false, FileName, Missing.Value);
        }

        private static int GetExcelColumnNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName))
                throw new ArgumentNullException("Invalid column name parameter");

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            char ch;
            for (int i = 0; i < columnName.Length; i++)
            {
                ch = columnName[i];

                if (char.IsDigit(ch))
                    throw new ArgumentNullException("Invalid column name parameter on character " + ch);

                sum *= 26;
                sum += (ch - 'A' + 1);
            }

            return sum;
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}
