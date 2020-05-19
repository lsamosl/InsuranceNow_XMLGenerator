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
        #region Constants Declaration

        int DRIVER_BIRTHDATE, DRIVER_DATEFIRSTLICENSE, DRIVER_DRIVERNUMBER, DRIVER_DRIVERSTATUS, DRIVER_FIRSTNAME, DRIVER_GENDER,
            DRIVER_LASTNAME, DRIVER_LICENSENUMBER, DRIVER_LICENSESTATE, DRIVER_MARITALSTATUS, DRIVER_MATUREDRIVER, DRIVER_MIDDLENAME,
            DRIVER_OCCUPATION, DRIVER_POLICY, DRIVER_RELATIONSHIP, DRIVER_SR22,
            VEHICLE_ANNUALMILEAGE, VEHICLE_COLLDEDUCT, VEHICLE_COMPDEDUCT, VEHICLE_CURRENTODOMETER, VEHICLE_DR, VEHICLE_LESSOR, VEHICLE_MAKE,
            VEHICLE_MODEL, VEHICLE_NO, VEHICLE_POLICY, VEHICLE_PUB, VEHICLE_RENTAL, VEHICLE_VIN;

        #endregion

        public string FileName { get; set; }
        public int IndexRow { get; set; }

        public ExcelUtil(string fileName)
        {
            FileName = fileName;
            IndexRow = 2;

            #region Constants initialization

            DRIVER_BIRTHDATE = GetExcelColumnNumber(ExcelConfiguration.Driver_birthDate);
            DRIVER_DATEFIRSTLICENSE = GetExcelColumnNumber(ExcelConfiguration.Driver_dateFirstLicense);
            DRIVER_DRIVERNUMBER = GetExcelColumnNumber(ExcelConfiguration.Driver_DriverNumber);
            DRIVER_DRIVERSTATUS = GetExcelColumnNumber(ExcelConfiguration.Driver_driverStatus);
            DRIVER_FIRSTNAME = GetExcelColumnNumber(ExcelConfiguration.Driver_FirstName);
            DRIVER_GENDER = GetExcelColumnNumber(ExcelConfiguration.Driver_Gender);
            DRIVER_LASTNAME = GetExcelColumnNumber(ExcelConfiguration.Driver_LastName);
            DRIVER_LICENSENUMBER = GetExcelColumnNumber(ExcelConfiguration.Driver_licenseNumber);
            DRIVER_LICENSESTATE = GetExcelColumnNumber(ExcelConfiguration.Driver_licenseState);
            DRIVER_MARITALSTATUS = GetExcelColumnNumber(ExcelConfiguration.Driver_maritalStatus);
            DRIVER_MATUREDRIVER = GetExcelColumnNumber(ExcelConfiguration.Driver_matureDriver);
            DRIVER_MIDDLENAME = GetExcelColumnNumber(ExcelConfiguration.Driver_MiddleName);
            DRIVER_OCCUPATION = GetExcelColumnNumber(ExcelConfiguration.Driver_occupation);
            DRIVER_POLICY = GetExcelColumnNumber(ExcelConfiguration.Driver_Policy);
            DRIVER_RELATIONSHIP = GetExcelColumnNumber(ExcelConfiguration.Driver_relationShip);
            DRIVER_SR22 = GetExcelColumnNumber(ExcelConfiguration.Driver_sr22);
            VEHICLE_ANNUALMILEAGE = GetExcelColumnNumber(ExcelConfiguration.Vehicle_annualMileage);
            VEHICLE_COLLDEDUCT = GetExcelColumnNumber(ExcelConfiguration.Vehicle_COLLDEDUCT);
            VEHICLE_COMPDEDUCT = GetExcelColumnNumber(ExcelConfiguration.Vehicle_COMPDEDUCT);
            VEHICLE_CURRENTODOMETER = GetExcelColumnNumber(ExcelConfiguration.Vehicle_currentOdometer);
            VEHICLE_DR = GetExcelColumnNumber(ExcelConfiguration.Vehicle_DR);
            VEHICLE_LESSOR = GetExcelColumnNumber(ExcelConfiguration.Vehicle_LESSOR);
            VEHICLE_MAKE = GetExcelColumnNumber(ExcelConfiguration.Vehicle_Make);
            VEHICLE_MODEL = GetExcelColumnNumber(ExcelConfiguration.Vehicle_Model);
            VEHICLE_NO = GetExcelColumnNumber(ExcelConfiguration.Vehicle_NO);
            VEHICLE_POLICY = GetExcelColumnNumber(ExcelConfiguration.Vehicle_Policy);
            VEHICLE_PUB = GetExcelColumnNumber(ExcelConfiguration.Vehicle_PUB);
            VEHICLE_RENTAL = GetExcelColumnNumber(ExcelConfiguration.Vehicle_Rental);
            VEHICLE_VIN = GetExcelColumnNumber(ExcelConfiguration.Vehicle_VIN);

            #endregion
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
                _Worksheet PolicySheet = xlWorkbook.Sheets[ExcelConfiguration.PolicySheet];
                _Worksheet DriverSheet = xlWorkbook.Sheets[ExcelConfiguration.DriverSheet];
                _Worksheet VehicleSheet = xlWorkbook.Sheets[ExcelConfiguration.VehicleSheet];
                _Worksheet AutoGeneralSheet = xlWorkbook.Sheets[ExcelConfiguration.AutoGeneralSheet];

                Range rangeObject = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_PolicyNumber];
                string policyNumber = (string)rangeObject.Value2;

                Policy policy = new Policy();

                do
                {
                    policy = new Policy {
                        General = new General()
                        {
                            MedCover = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_medCover].Value2.ToString(),
                            BiCover = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_biCover].Value2.ToString(),
                            Umbi = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_umbi].Value2.ToString(),
                            UmpdCdw = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_umpdcdw].Value2.ToString(),
                            LimMex = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_limMex].Value2.ToString(),
                            RoadAssis = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_roadAssis].Value2.ToString(),
                            PdCover = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_pdCover].Value2.ToString()
                        },

                        FirstName = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_firstName].Value2.ToString(),
                        LastName = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_lastName].Value2.ToString(),
                        BillReminder = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_BillReminder].Value2.ToString(),
                        BirthDate = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_birthDate].Value2.ToString(),
                        PrimaryPhone = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_primaryPhone].Value2.ToString(),
                        MailingAddress = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_mailingAddress].Value2.ToString(),
                        MailingCity = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_mailingCity].Value2.ToString(),
                        MailingState = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_mailingState].Value2.ToString(),
                        MailingZip = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_mailingZip].Value2.ToString(),
                        GaragingAddress = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_garagingAddress].Value2.ToString(),
                        GaragingCity = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_garagingCity].Value2.ToString(),
                        GaragingState = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_garagingState].Value2.ToString(),
                        GaragingZip = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_garagingZip].Value2.ToString(),
                        InceptionDate = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_inceptionDate].Value2.ToString(),
                        EffectiveDate = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_effectiveDate].Value2.ToString(),
                        ExpirationDate = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_expirationDate].Value2.ToString(),
                        Term = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_term].Value2.ToString(),
                        ProducerCode = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_producerCode].Value2.ToString(),
                        PolicyNumber = policyNumber,

                        Drivers = new List<Driver>(),
                        Vehicles = new List<Vehicle>()
                    };

                    Range xlrangeDriver = DriverSheet.UsedRange;
                    xlrangeDriver.AutoFilter(GetExcelColumnNumber(ExcelConfiguration.Driver_Policy), policyNumber, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                    Range filteredRangeDriver = xlrangeDriver.SpecialCells(XlCellType.xlCellTypeVisible, XlSpecialCellsValue.xlTextValues);

                    for (int areaId = 1; areaId <= filteredRangeDriver.Areas.Count; areaId++)
                    {
                        Range areaRange = filteredRangeDriver.Areas.get_Item(areaId);
                        object[,] areaValues = areaRange.Value2;

                        for (int i = 2; i <= areaValues.GetLength(0); i++)
                        {
                            policy.Drivers.Add(new Driver
                            {
                                Gender = areaValues[i, DRIVER_GENDER] != null ? areaValues[i, DRIVER_GENDER].ToString() : string.Empty,
                                BirthDate = areaValues[i, DRIVER_BIRTHDATE] != null? areaValues[i, DRIVER_BIRTHDATE].ToString() : string.Empty,
                                MaritalStatus = areaValues[i, DRIVER_MARITALSTATUS] != null? areaValues[i, DRIVER_MARITALSTATUS].ToString() : string.Empty,
                                Occupation = areaValues[i, DRIVER_OCCUPATION] != null? areaValues[i, DRIVER_OCCUPATION].ToString() : string.Empty,
                                DriverNumber =areaValues[i, DRIVER_DRIVERNUMBER] != null ? Int32.Parse(areaValues[i, DRIVER_DRIVERNUMBER].ToString()) : 0,
                                DriverStatus = areaValues[i, DRIVER_DRIVERSTATUS] != null? areaValues[i, DRIVER_DRIVERSTATUS].ToString() : string.Empty,
                                LicenseNumber = areaValues[i, DRIVER_LICENSENUMBER] != null? areaValues[i, DRIVER_LICENSENUMBER].ToString() : string.Empty,
                                DateFirstLicense = areaValues[i, DRIVER_DATEFIRSTLICENSE] != null? areaValues[i, DRIVER_DATEFIRSTLICENSE].ToString() : string.Empty,
                                LicenseState = areaValues[i, DRIVER_LICENSESTATE] != null? areaValues[i, DRIVER_LICENSESTATE].ToString() : string.Empty,
                                RelationShip = areaValues[i, DRIVER_RELATIONSHIP] != null? areaValues[i, DRIVER_RELATIONSHIP].ToString() : string.Empty,
                                MatureDriver = areaValues[i, DRIVER_MATUREDRIVER] != null? areaValues[i, DRIVER_MATUREDRIVER].ToString() : string.Empty,
                                Sr22 = areaValues[i, DRIVER_SR22] != null ? areaValues[i, DRIVER_SR22].ToString() : string.Empty,
                            });
                        }
                    }

                    Range xlrangeVehicle = VehicleSheet.UsedRange;
                    xlrangeVehicle.AutoFilter(GetExcelColumnNumber(ExcelConfiguration.Vehicle_Policy), policyNumber, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                    Range filteredRangeVehicle = xlrangeVehicle.SpecialCells(XlCellType.xlCellTypeVisible, XlSpecialCellsValue.xlTextValues);

                    for (int areaId = 1; areaId <= filteredRangeVehicle.Areas.Count; areaId++)
                    {
                        Range areaRange = filteredRangeVehicle.Areas.get_Item(areaId);
                        object[,] areaValues = areaRange.Value2;

                        for (int i = 2; i <= areaValues.GetLength(0); i++)
                        {
                            policy.Vehicles.Add(new Vehicle
                            {
                                VIN = areaValues[i, VEHICLE_VIN] != null ? areaValues[i, VEHICLE_VIN].ToString() : string.Empty,
                                CollDeduct = areaValues[i, VEHICLE_COLLDEDUCT] != null ? areaValues[i, VEHICLE_COLLDEDUCT].ToString() : string.Empty,
                                CompDeduct = areaValues[i, VEHICLE_COMPDEDUCT] != null ? areaValues[i, VEHICLE_COMPDEDUCT].ToString(): string.Empty,
                                AnnualMileage = areaValues[i, VEHICLE_ANNUALMILEAGE] != null ? areaValues[i, VEHICLE_ANNUALMILEAGE].ToString(): string.Empty,
                                CurrentOdometer = areaValues[i, VEHICLE_CURRENTODOMETER] != null ? areaValues[i, VEHICLE_CURRENTODOMETER].ToString(): string.Empty,
                                Lessor = areaValues[i, VEHICLE_LESSOR] != null ? areaValues[i, VEHICLE_LESSOR].ToString(): string.Empty,
                                Rental = areaValues[i, VEHICLE_RENTAL] != null ? areaValues[i, VEHICLE_RENTAL].ToString(): string.Empty,
                                No = areaValues[i, VEHICLE_NO] != null ? areaValues[i, VEHICLE_NO].ToString(): string.Empty,
                                Pub = areaValues[i, VEHICLE_PUB] != null ? areaValues[i, VEHICLE_PUB].ToString(): string.Empty,
                                Dr = areaValues[i, VEHICLE_DR] != null ? areaValues[i, VEHICLE_DR].ToString(): string.Empty,
                            });
                        }
                    }

                    Policies.Add(policy);

                    IndexRow++;
                    rangeObject = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_PolicyNumber];
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
    }
}
