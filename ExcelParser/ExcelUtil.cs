using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Models;
using Microsoft.Office.Interop.Excel;
using Configurations;

namespace ExcelParser
{
    public class ExcelUtil
    {
        public string FileName { get; set; }
        public int IndexRow { get; set; }

        public ExcelUtil(string fileName)
        {
            FileName = fileName;
            IndexRow = 2;
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
                _Worksheet PaymentSheet = xlWorkbook.Sheets[ExcelConfiguration.PaymentSheet];

                Range rangeObject;
                string policyNumber = string.Empty;
                string dateFormat = "yyyyMMdd";

                Policy policy = new Policy();
                List<General> GeneralList = new List<General>();
                List<Driver> DriverList = new List<Driver>();
                List<Vehicle> VehicleList = new List<Vehicle>();

                int term;
                DateTime effectiveDate, expirationDate, birthDate;

                /*General*/
                rangeObject = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_PolicyNumber];
                policyNumber = (string)rangeObject.Value2;

                do
                {
                    General general = new General()
                    {
                        MedCover = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_medCover].Value2.ToString(),
                        BiCover = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_biCover].Value2.ToString(),
                        Umbi = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_umbi].Value2.ToString(),
                        UmpdCdw = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_umpdcdw].Value2.ToString(),
                        LimMex = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_limMex].Value2.ToString(),
                        RoadAssis = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_roadAssis].Value2.ToString(),
                        PdCover = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_pdCover].Value2.ToString(),
                        PolicyNumber = policyNumber
                    };

                    general.Coverages = new GeneralCoverages(general.BiCover, general.PdCover, general.Umbi, general.MedCover);
                    general.Coverages.ProcessCoverages();

                    GeneralList.Add(general);

                    IndexRow++;
                    rangeObject = AutoGeneralSheet.Cells[IndexRow, ExcelConfiguration.AutoGeneral_PolicyNumber];
                    policyNumber = (string)rangeObject.Value2;
                }
                while (!string.IsNullOrEmpty(policyNumber));

                /*Drivers*/
                IndexRow = 2;
                rangeObject = DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_PolicyNumber];
                policyNumber = (string)rangeObject.Value2;

                do
                {
                    DriverList.Add(new Driver
                    {
                        FirstName = DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_FirstName].Value2.ToString().Trim(),
                        MiddleName = DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_MiddleName].Value2.ToString().Trim(),
                        LastName = DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_LastName].Value2.ToString().Trim(),
                        Gender = DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_Gender].Value2.ToString().Trim() == "M" ? "Male" : "Female",
                        BirthDate = DateTime.FromOADate(DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_birthDate].Value2).ToString(dateFormat),
                        MaritalStatus = DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_maritalStatus].Value2.ToString().Trim() == "M" ? "Married" : "Single",
                        Occupation = DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_occupation].Value2.ToString().Trim() == "EMPLOYED" ? "Employed" : DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_occupation].Value2.ToString().Trim() ||
                                     DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_occupation].Value2.ToString().Trim() == "" ? "Employed" : DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_occupation].Value2.ToString().Trim() ||
                                     DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_occupation].Value2.ToString().Trim() == "-1" ? "Employed" : DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_occupation].Value2.ToString().Trim(),
                        DriverNumber = DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_DriverNumber].Value2.ToString(),
                        DriverStatus = DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_driverStatus].Value2.ToString().Trim(),
                        LicenseNumber = DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_licenseNumber].Value2.ToString().Trim(),
                        DateFirstLicense = DateTime.FromOADate(DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_dateFirstLicense].Value2).ToString(dateFormat),
                        LicenseState = DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_licenseState].Value2.ToString(),
                        RelationShip = DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_relationShip].Value2.ToString().Trim() == "Insured" ? "Self" : DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_relationShip].Value2.ToString().Trim(),
                        MatureDriver = DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_matureDriver].Value2.ToString().Trim() == "Y" ? "Yes" : "No",
                        Sr22 = DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_sr22].Value2.ToString().Trim() == "Y" ? "Yes" : "No",
                        PolicyNumber = policyNumber
                    });

                    IndexRow++;
                    rangeObject = DriverSheet.Cells[IndexRow, ExcelConfiguration.Driver_PolicyNumber];
                    policyNumber = (string)rangeObject.Value2;
                }
                while (!string.IsNullOrEmpty(policyNumber));


                /*Vehicles*/
                IndexRow = 2;
                rangeObject = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_PolicyNumber];
                policyNumber = (string)rangeObject.Value2;

                do
                {
                    Vehicle vehicle = new Vehicle
                    {
                        VIN = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_VIN].Value2.ToString().Trim(),
                        CollDeduct = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_COLLDEDUCT].Value2.ToString(),
                        CompDeduct = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_COMPDEDUCT].Value2.ToString(),
                        AnnualMileage = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_annualMileage].Value2.ToString(),
                        CurrentOdometer = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_currentOdometer].Value2.ToString(),
                        Lessor = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_LESSOR].Value2.ToString(),
                        Rental = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_Rental].Value2.ToString(),
                        No = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_NO].Value2.ToString(),
                        Pub = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_PUB].Value2.ToString(),
                        Dr = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_DR].Value2.ToString(),
                        Make = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_Make].Value2.ToString().Trim(),
                        Model = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_Model].Value2.ToString().Trim(),
                        PolicyNumber = policyNumber
                    };

                    vehicle.Coverages = new VehicleCoverages(vehicle.CollDeduct, vehicle.CompDeduct, vehicle.Rental);
                    vehicle.Coverages.ProcessCoverages();

                    VehicleList.Add(vehicle);

                    IndexRow++;
                    rangeObject = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_PolicyNumber];
                    policyNumber = (string)rangeObject.Value2;
                }
                while (!string.IsNullOrEmpty(policyNumber));

                /*Payment plans*/
                IndexRow = 2;
                rangeObject = PaymentSheet.Cells[IndexRow, ExcelConfiguration.Payment_PolicyNumber];
                policyNumber = (string)rangeObject.Value2;

                do
                {
                    Pay vehicle = new Vehicle
                    {
                        VIN = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_VIN].Value2.ToString().Trim(),
                        CollDeduct = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_COLLDEDUCT].Value2.ToString(),
                        CompDeduct = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_COMPDEDUCT].Value2.ToString(),
                        AnnualMileage = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_annualMileage].Value2.ToString(),
                        CurrentOdometer = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_currentOdometer].Value2.ToString(),
                        Lessor = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_LESSOR].Value2.ToString(),
                        Rental = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_Rental].Value2.ToString(),
                        No = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_NO].Value2.ToString(),
                        Pub = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_PUB].Value2.ToString(),
                        Dr = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_DR].Value2.ToString(),
                        Make = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_Make].Value2.ToString().Trim(),
                        Model = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_Model].Value2.ToString().Trim(),
                        PolicyNumber = policyNumber
                    };

                    vehicle.Coverages = new VehicleCoverages(vehicle.CollDeduct, vehicle.CompDeduct, vehicle.Rental);
                    vehicle.Coverages.ProcessCoverages();

                    VehicleList.Add(vehicle);

                    IndexRow++;
                    rangeObject = VehicleSheet.Cells[IndexRow, ExcelConfiguration.Vehicle_PolicyNumber];
                    policyNumber = (string)rangeObject.Value2;
                }
                while (!string.IsNullOrEmpty(policyNumber));

                /*Policies*/
                IndexRow = 2;
                rangeObject = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_PolicyNumber];
                policyNumber = (string)rangeObject.Value2;

                do
                {
                    term = Int32.Parse(PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_term].Value2.ToString());
                    effectiveDate = DateTime.FromOADate(PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_effectiveDate].Value2);
                    birthDate = DateTime.FromOADate(PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_birthDate].Value2);
                    expirationDate = effectiveDate.AddMonths(term);

                    policy = new Policy
                    {
                        FirstName = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_firstName].Value2.ToString().Trim(),
                        LastName = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_lastName].Value2.ToString().Trim(),
                        BillReminder = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_BillReminder].Value2.ToString(),
                        BirthDate = DateTime.FromOADate(PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_birthDate].Value2).ToString(dateFormat),
                        PrimaryPhone = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_primaryPhone].Value2.ToString(),
                        MailingAddress = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_mailingAddress].Value2.ToString().Trim(),
                        MailingCity = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_mailingCity].Value2.ToString().Trim(),
                        MailingState = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_mailingState].Value2.ToString().Trim(),
                        MailingZip = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_mailingZip].Value2.ToString().Trim(),
                        GaragingAddress = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_garagingAddress].Value2.ToString().Trim(),
                        GaragingCity = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_garagingCity].Value2.ToString().Trim(),
                        GaragingState = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_garagingState].Value2.ToString().Trim(),
                        GaragingZip = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_garagingZip].Value2.ToString().Trim(),
                        InceptionDate = DateTime.FromOADate(PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_inceptionDate].Value2).ToString(dateFormat),
                        EffectiveDate = effectiveDate.ToString(dateFormat),
                        ExpirationDate = expirationDate.ToString(dateFormat),
                        Term = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_term].Value2.ToString(),
                        ProducerCode = PolicySheet.Cells[IndexRow, ExcelConfiguration.Policy_producerCode].Value2.ToString().Replace(" ", string.Empty),
                        Age = (DateTime.Today.Year - birthDate.Year).ToString(),
                        PolicyNumber = policyNumber,

                        General = GeneralList.Where(x => x.PolicyNumber.Equals(policyNumber)).FirstOrDefault(),
                        Drivers = DriverList.Where(x => x.PolicyNumber.Equals(policyNumber)).ToList(),
                        Vehicles = VehicleList.Where(x => x.PolicyNumber.Equals(policyNumber)).ToList()
                    };

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
    }
}
