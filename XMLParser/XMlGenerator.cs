using System;
using System.Collections.Generic;
using System.Xml;
using Services;
using System.Net;
using Models;
using Configurations;

namespace XMLParser
{
    public class XMLGenerator
    {
        public Policy Policy { get; set; }
        public XMLGenerator(string path, Policy _Policy)
        {
            Policy = _Policy;
            Path = path;
        }

        public string Path { get; set; }

        public void Generate()
        {
            try
            {
                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Indent = true;
                settings.OmitXmlDeclaration = true;

                int countVehicles = 1;
                int countDrivers = 1;
                Driver principalDriver = Policy.Drivers.Count > 0 ? Policy.Drivers[0] : null;
                string fullName = string.Format("{0} {1}", Policy.FirstName, Policy.LastName);
                string Gender = principalDriver != null ? (principalDriver.Gender == "M" ? "Male" : "Female") : string.Empty;
                string term = string.Format("{0} {1}", Int32.Parse(Policy.Term).ToString(), "Months");
                string[] NameTypeCd = XMLStaticValues.PartyInfo_NameInfo_NameTypeCd.Split('|');
                string[] PersonTypeCd = XMLStaticValues.PartyInfo_PersonInfo_PersonTypeCd.Split('|'); 
                string[] AddrTypeCd = XMLStaticValues.PartyInfo_Addr_AddrTypeCd.Split('|');
                string[] QuestionReply_Names = XMLStaticValues.QuestionReplies_QuestionReply_Name.Split('|');
                string[] biCoverArray = Policy.General.BiCover.Split('/'); 
                string vehicleDescription = string.Empty;
                string[] limits = XMLStaticValues.DTOCoverage1_DTOLimit_LimitCd.Split('|');
                string[] limits_values = XMLStaticValues.DTOCoverage1_DTOLimit_TypeCd.Split('|');

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                System.Net.ServicePointManager.ServerCertificateValidationCallback += (se, cert, chain, sslerror) => true;
                Tokenizer tokenizerService = new Tokenizer();
                tokenizerService.PolicyNumber = Policy.PolicyNumber;

                using (XmlWriter writer = XmlWriter.Create(Path, settings))
                {
                    writer.WriteStartDocument();

                    writer.WriteStartElement("DTORoot");

                    #region <DTOApplication>
                    writer.WriteStartElement("DTOApplication");

                    writer.WriteAttributeString("Version", XMLStaticValues.DTORoot_DTOApplication_Version);
                    writer.WriteAttributeString("Status", XMLStaticValues.DTORoot_DTOApplication_Status);
                    writer.WriteAttributeString("TypeCd", XMLStaticValues.DTORoot_DTOApplication_TypeCd);
                    writer.WriteAttributeString("Description", Policy.PolicyNumber);
                    writer.WriteAttributeString("ReadyToRateInd", XMLStaticValues.DTORoot_DTOApplication_ReadyToRateInd);                                    

                    #region <QuestionReplies>
                    writer.WriteStartElement("QuestionReplies");
                    writer.WriteAttributeString("QuestionSourceMDA", XMLStaticValues.DTOApplication_QuestionReplies_QuestionSourceMDA);

                    foreach (string q in QuestionReply_Names)
                    {
                        StartElementWithAttributes(writer, "QuestionReply",
                           new Dictionary<string, string> {
                               { "Name", q } ,
                               { "Value", XMLStaticValues.QuestionReplies_QuestionReply_Value } ,
                               { "VisibleInd", XMLStaticValues.QuestionReplies_QuestionReply_VisibleInd } ,
                           });
                    }

                    writer.WriteEndElement(); // End QuestionReplies
                    #endregion

                    #region <DTOLine>
                    writer.WriteStartElement("DTOLine");

                    writer.WriteAttributeString("StatusCd", XMLStaticValues.DTOApplication_DTOLine_StatusCd);
                    writer.WriteAttributeString("LineCd", XMLStaticValues.DTOApplication_DTOLine_LineCd);
                    writer.WriteAttributeString("RatingService", XMLStaticValues.DTOApplication_DTOLine_RatingService);

                    if (Policy.General.Coverages.MED.hasCoverage)
                        writer.WriteAttributeString("MedPayLimit", Policy.General.Coverages.MED.InputValue);
                    else
                        writer.WriteAttributeString("MedPayLimit", "None");

                    if (Policy.General.Coverages.BI.hasCoverage)
                        writer.WriteAttributeString("BILimit", Policy.General.Coverages.BI.InputValue);
                    else
                        writer.WriteAttributeString("BILimit", "None");

                    if (Policy.General.Coverages.PD.hasCoverage)
                        writer.WriteAttributeString("PDLimit", Policy.General.Coverages.PD.InputValue);
                    else
                        writer.WriteAttributeString("PDLimit", "None");

                    if (Policy.General.Coverages.UMBI.hasCoverage)
                        writer.WriteAttributeString("UMBILimit", Policy.General.Coverages.UMBI.InputValue);
                    else
                        writer.WriteAttributeString("UMBILimit", "None");

                    writer.WriteAttributeString("UMPDWCDInd", !String.IsNullOrEmpty(Policy.General.UmpdCdw) ? Policy.General.UmpdCdw : "No");
                    writer.WriteAttributeString("LimitedMexicoInd", !String.IsNullOrEmpty(Policy.General.LimMex) ? Policy.General.LimMex : "No");
                    writer.WriteAttributeString("RoadAssistInd", !String.IsNullOrEmpty(Policy.General.RoadAssis) ? Policy.General.RoadAssis : "No");
                    writer.WriteAttributeString("PolicyType", XMLStaticValues.DTOApplication_DTOLine_PolicyType);
                    writer.WriteAttributeString("TotalNumVehicles", Policy.Vehicles.Count.ToString());
                    writer.WriteAttributeString("TotalNumDrivers", Policy.Drivers.Count.ToString());

                    #region <DTORisk>
                    foreach (Vehicle v in Policy.Vehicles)
                    {
                        writer.WriteStartElement("DTORisk");

                        writer.WriteAttributeString("TypeCd", XMLStaticValues.DTOLine_DTORisk_TypeCd);
                        writer.WriteAttributeString("Status", XMLStaticValues.DTOLine_DTORisk_Status);
                        writer.WriteAttributeString("Description", string.Format("{0} {1} - {2}", v.Make, v.Model, v.VIN));
                        writer.WriteAttributeString("RiskBeanName", XMLStaticValues.DTOLine_DTORisk_RiskBeanName);
                        writer.WriteAttributeString("RiskAddDt", Policy.EffectiveDate); 
                        writer.WriteAttributeString("RiskAddPolicyVersion", XMLStaticValues.DTOLine_DTORisk_RiskAddPolicyVersion);
                        writer.WriteAttributeString("RiskAddTransactionNo", XMLStaticValues.DTOLine_DTORisk_RiskAddTransactionNo);

                        if (Policy.General.Coverages.BI.hasCoverage)
                        {
                            writer.WriteStartElement("DTOCoverage");

                            writer.WriteAttributeString("Status", XMLStaticValues.Coverage_Status);
                            writer.WriteAttributeString("CoverageCd", XMLStaticValues.Coverage_BI_Name);
                            writer.WriteAttributeString("Description", XMLStaticValues.Coverage_BI_Description);

                            writer.WriteStartElement("DTOLimit");
                            writer.WriteAttributeString("LimitCd", limits[0]);
                            writer.WriteAttributeString("TypeCd", limits_values[0]);
                            writer.WriteAttributeString("Value", Policy.General.Coverages.BI.Value1);
                            writer.WriteEndElement(); //End DTOLimit

                            writer.WriteStartElement("DTOLimit");
                            writer.WriteAttributeString("LimitCd", limits[1]);
                            writer.WriteAttributeString("TypeCd", limits_values[1]);
                            writer.WriteAttributeString("Value", Policy.General.Coverages.BI.Value2);
                            writer.WriteEndElement(); //End DTOLimit

                            writer.WriteEndElement(); //End DTOCoverage
                        }

                        writer.WriteStartElement("DTOVehicle");

                        writer.WriteAttributeString("VehNumber", countVehicles.ToString());
                        writer.WriteAttributeString("Status", XMLStaticValues.DTORisk_DTOVehicle_Status);
                        writer.WriteAttributeString("VehIdentificationNumber", v.VIN);
                        writer.WriteAttributeString("ValidVinInd", XMLStaticValues.DTORisk_DTOVehicle_ValidVinInd);
                        writer.WriteAttributeString("VehUseCd", XMLStaticValues.DTORisk_DTOVehicle_VehUseCd);
                        writer.WriteAttributeString("CollisionDed", v.Coverages.COLL.hasCoverage ? v.Coverages.COLL.InputValue : "No Coverage");
                        writer.WriteAttributeString("ComprehensiveDed", v.Coverages.CPR.hasCoverage ? v.Coverages.CPR.InputValue : "No Coverage");
                        writer.WriteAttributeString("EstimatedAnnualDistance", v.AnnualMileage);
                        writer.WriteAttributeString("OdometerReading", v.CurrentOdometer);
                        writer.WriteAttributeString("LessorLiabilityInd", !string.IsNullOrEmpty(v.Lessor) ? v.Lessor : "No");
                        writer.WriteAttributeString("RentalReimbursement", !string.IsNullOrEmpty(v.Rental) ? v.Rental : "None");
                        writer.WriteAttributeString("DamageInd", XMLStaticValues.DTORisk_DTOVehicle_DamageInd);
                        writer.WriteAttributeString("RegisterdOwner", XMLStaticValues.DTORisk_DTOVehicle_RegisterdOwner);
                        writer.WriteAttributeString("NonOwnedVehicleInd", !string.IsNullOrEmpty(v.No) ? v.No : "No");
                        writer.WriteAttributeString("PermissiveUseInd", !string.IsNullOrEmpty(v.Pub) ? v.Pub : "No");
                        writer.WriteAttributeString("DirectRepairInd", !string.IsNullOrEmpty(v.Dr) ? v.Dr : "No");

                        writer.WriteEndElement(); //End DTOVehicle

                        countVehicles++;

                        if (Policy.General.Coverages.PD.hasCoverage)
                        {
                            writer.WriteStartElement("DTOCoverage");

                            writer.WriteAttributeString("Status", XMLStaticValues.Coverage_Status);
                            writer.WriteAttributeString("CoverageCd", XMLStaticValues.Coverage_PD_Name);
                            writer.WriteAttributeString("Description", XMLStaticValues.Coverage_PD_Description);

                            writer.WriteStartElement("DTOLimit");
                            writer.WriteAttributeString("LimitCd", limits[0]);
                            writer.WriteAttributeString("TypeCd", limits_values[0]);
                            writer.WriteAttributeString("Value", Policy.General.Coverages.PD.InputValue);
                            writer.WriteEndElement(); //End DTOLimit

                            writer.WriteEndElement(); //End DTOCoverage
                        }

                        if (Policy.General.Coverages.MED.hasCoverage)
                        {
                            writer.WriteStartElement("DTOCoverage");

                            writer.WriteAttributeString("Status", XMLStaticValues.Coverage_Status);
                            writer.WriteAttributeString("CoverageCd", XMLStaticValues.Coverage_MED_Name);
                            writer.WriteAttributeString("Description", XMLStaticValues.Coverage_MED_Description);

                            writer.WriteStartElement("DTOLimit");
                            writer.WriteAttributeString("LimitCd", limits[0]);
                            writer.WriteAttributeString("TypeCd", limits_values[0]);
                            writer.WriteAttributeString("Value", Policy.General.Coverages.MED.InputValue);
                            writer.WriteEndElement(); //End DTOLimit

                            writer.WriteEndElement(); //End DTOCoverage
                        }

                        if (Policy.General.Coverages.UMBI.hasCoverage)
                        {
                            writer.WriteStartElement("DTOCoverage");

                            writer.WriteAttributeString("Status", XMLStaticValues.Coverage_Status);
                            writer.WriteAttributeString("CoverageCd", XMLStaticValues.Coverage_UMBI_Name);
                            writer.WriteAttributeString("Description", XMLStaticValues.Coverage_UMBI_Description);

                            writer.WriteStartElement("DTOLimit");
                            writer.WriteAttributeString("LimitCd", limits[0]);
                            writer.WriteAttributeString("TypeCd", limits_values[0]);
                            writer.WriteAttributeString("Value", Policy.General.Coverages.UMBI.Value1);
                            writer.WriteEndElement(); //End DTOLimit

                            writer.WriteStartElement("DTOLimit");
                            writer.WriteAttributeString("LimitCd", limits[1]);
                            writer.WriteAttributeString("TypeCd", limits_values[1]);
                            writer.WriteAttributeString("Value", Policy.General.Coverages.UMBI.Value2);
                            writer.WriteEndElement(); //End DTOLimit

                            writer.WriteEndElement(); //End DTOCoverage
                        }

                        if (v.Coverages.CPR.hasCoverage)
                        {
                            writer.WriteStartElement("DTOCoverage");

                            writer.WriteAttributeString("Status", XMLStaticValues.Coverage_Status);
                            writer.WriteAttributeString("CoverageCd", XMLStaticValues.Coverage_CPR_Name);
                            writer.WriteAttributeString("Description", XMLStaticValues.Coverage_CPR_Description);

                            writer.WriteStartElement("DTODeductible");

                            writer.WriteAttributeString("DeductibleCd", XMLStaticValues.Deductible_DeductibleCd);
                            writer.WriteAttributeString("TypeCd", XMLStaticValues.Deductible_TypeCd);
                            writer.WriteAttributeString("Value", v.Coverages.CPR.InputValue);

                            writer.WriteEndElement(); //End DTODeductible

                            writer.WriteEndElement(); //End DTOCoverage
                        }

                        if (v.Coverages.COLL.hasCoverage)
                        {
                            writer.WriteStartElement("DTOCoverage");

                            writer.WriteAttributeString("Status", XMLStaticValues.Coverage_Status);
                            writer.WriteAttributeString("CoverageCd", XMLStaticValues.Coverage_COLL_Name);
                            writer.WriteAttributeString("Description", XMLStaticValues.Coverage_COLL_Description);

                            writer.WriteStartElement("DTODeductible");

                            writer.WriteAttributeString("DeductibleCd", XMLStaticValues.Deductible_DeductibleCd);
                            writer.WriteAttributeString("TypeCd", XMLStaticValues.Deductible_TypeCd);
                            writer.WriteAttributeString("Value", v.Coverages.COLL.InputValue);

                            writer.WriteEndElement(); //End DTODeductible

                            writer.WriteEndElement(); //End DTOCoverage
                        }

                        if (v.Coverages.Rental.hasCoverage)
                        {
                            writer.WriteStartElement("DTOCoverage");

                            writer.WriteAttributeString("Status", XMLStaticValues.Coverage_Status);
                            writer.WriteAttributeString("CoverageCd", XMLStaticValues.Coverage_Rental_Name);
                            writer.WriteAttributeString("Description", XMLStaticValues.Coverage_Rental_Description);

                            writer.WriteStartElement("DTOLimit");
                            writer.WriteAttributeString("LimitCd", limits[0]);
                            writer.WriteAttributeString("TypeCd", limits_values[0]);
                            writer.WriteAttributeString("Value", v.Coverages.Rental.Value1);
                            writer.WriteEndElement(); //End DTOLimit

                            writer.WriteStartElement("DTOLimit");
                            writer.WriteAttributeString("LimitCd", limits[1]);
                            writer.WriteAttributeString("TypeCd", limits_values[1]);
                            writer.WriteAttributeString("Value", v.Coverages.Rental.Value2);
                            writer.WriteEndElement(); //End DTOLimit

                            writer.WriteEndElement(); //End DTOCoverage
                        }

                        writer.WriteEndElement(); //End DTORisk
                    }
                    #endregion

                    writer.WriteEndElement(); //End DTOLine
                    #endregion

                    #region <PartyInfo>
                    writer.WriteStartElement("PartyInfo");

                    writer.WriteAttributeString("PartyTypeCd", XMLStaticValues.DTOApplication_PartyInfo_PartyTypeCd);
                    writer.WriteAttributeString("Status", XMLStaticValues.DTOApplication_PartyInfo_Status);

                    #region <PersonInfo>
                    writer.WriteStartElement("PersonInfo");

                    writer.WriteAttributeString("PersonTypeCd", PersonTypeCd[0]);
                    writer.WriteAttributeString("GenderCd", Gender);
                    writer.WriteAttributeString("BirthDt", principalDriver != null ? principalDriver.BirthDate : string.Empty);
                    writer.WriteAttributeString("MaritalStatusCd", principalDriver != null ? principalDriver.MaritalStatus : string.Empty);
                    writer.WriteAttributeString("OccupationClassCd", principalDriver != null ? principalDriver.Occupation : string.Empty);
                    writer.WriteAttributeString("Age", Policy.Age); //TODO

                    writer.WriteEndElement(); //End PersonInfo
                    #endregion

                    #region <DriverInfo>

                    Driver d = Policy.Drivers[0];
                    tokenizerService.DriverName = string.Format("{0} {1}", d.FirstName, d.LastName);
                    var response = tokenizerService.Detokenize(d.LicenseNumber);
                    writer.WriteStartElement("DriverInfo");

                    writer.WriteAttributeString("DriverInfoCd", XMLStaticValues.PartyInfo_DriverInfo_DriverInfoCd);
                    writer.WriteAttributeString("DriverNumber", 1.ToString());
                    writer.WriteAttributeString("DriverStatusCd", d.DriverStatus == "Active" ? "Principal" : "Excluded");

                    if (response.Result)
                        writer.WriteAttributeString("LicenseNumber", response.DetokenizedString);
                    else
                        writer.WriteAttributeString("LicenseNumber", string.Empty);

                    writer.WriteAttributeString("LicenseDt", d.DateFirstLicense);
                    writer.WriteAttributeString("LicensedStateProvCd", d.LicenseState);
                    writer.WriteAttributeString("RelationshipToInsuredCd", d.RelationShip);
                    writer.WriteAttributeString("MatureDriverInd", !String.IsNullOrEmpty(d.MatureDriver) ? d.MatureDriver : "No");
                    writer.WriteAttributeString("Race", XMLStaticValues.PartyInfo_DriverInfo_Race);
                    writer.WriteAttributeString("LicenseType", XMLStaticValues.PartyInfo_DriverInfo_LicenseType);
                    writer.WriteAttributeString("LicenseStatus", XMLStaticValues.PartyInfo_DriverInfo_LicenseStatus);
                    writer.WriteAttributeString("SR22Ind", !String.IsNullOrEmpty(d.Sr22) ? d.Sr22 : "No");
                    writer.WriteAttributeString("OriginalLicenseDt", d.DateFirstLicense);

                    countDrivers++;

                    writer.WriteEndElement(); //End DriverInfo
                                       
                    #endregion

                    #region <NameInfo>

                    writer.WriteStartElement("NameInfo");

                    writer.WriteAttributeString("NameTypeCd", NameTypeCd[0]);
                    writer.WriteAttributeString("GivenName", Policy.FirstName);
                    writer.WriteAttributeString("Surname", Policy.LastName);

                    writer.WriteEndElement(); //EndNameInfo
                    #endregion

                    writer.WriteEndElement(); //End PartyInfo
                    #endregion

                    #region <DTOTransactionInfo>
                    writer.WriteStartElement("DTOTransactionInfo");

                    writer.WriteAttributeString("TransactionCd", XMLStaticValues.DTOApplication_DTOTransactionInfo_TransactionCd);
                    writer.WriteAttributeString("TransactionEffectiveDt", Policy.EffectiveDate); //TODO
                    writer.WriteAttributeString("TransactionShortDescription", XMLStaticValues.DTOApplication_DTOTransactionInfo_TransactionShortDescription);
                    writer.WriteAttributeString("RewriteToProductVersion", XMLStaticValues.DTOApplication_DTOTransactionInfo_RewriteToProductVersion);
                    writer.WriteAttributeString("SourceCd", XMLStaticValues.DTOApplication_DTOTransactionInfo_SourceCd);
                    writer.WriteAttributeString("ChargeEndorsementFeeInd", XMLStaticValues.DTOApplication_DTOTransactionInfo_ChargeEndorsementFeeInd);

                    writer.WriteEndElement(); //End DTOTransactionInfo
                    #endregion

                    #region <DTOInsured>
                    writer.WriteStartElement("DTOInsured");

                    writer.WriteAttributeString("IndexName", fullName);
                    writer.WriteAttributeString("EntityTypeCd", XMLStaticValues.DTOApplication_DTOInsured_EntityTypeCd);
                    writer.WriteAttributeString("PreferredDeliveryMethod", XMLStaticValues.DTOApplication_DTOInsured_PreferredDeliveryMethod);
                    writer.WriteAttributeString("BillRemind", !string.IsNullOrEmpty(Policy.BillReminder) ? Policy.BillReminder : "No");

                    #region <PartyInfo>
                    writer.WriteStartElement("PartyInfo");

                    writer.WriteAttributeString("PartyTypeCd", XMLStaticValues.DTOInsured_PartyInfo_PartyTypeCd);

                    #region <PersonInfo>
                    writer.WriteStartElement("PersonInfo");

                    writer.WriteAttributeString("PersonTypeCd", PersonTypeCd[1]);
                    writer.WriteAttributeString("BirthDt", Policy.BirthDate);

                    writer.WriteEndElement(); //End PersonInfo
                    #endregion

                    #region <PhoneInfo>
                    writer.WriteStartElement("PhoneInfo");

                    writer.WriteAttributeString("PhoneTypeCd", XMLStaticValues.PartyInfo_PhoneInfo_PhoneTypeCd);
                    writer.WriteAttributeString("PhoneNumber", Policy.PrimaryPhone);
                    writer.WriteAttributeString("PhoneName", XMLStaticValues.PartyInfo_PhoneInfo_PhoneName);

                    writer.WriteEndElement(); //End PhoneInfo
                    #endregion

                    #region <NameInfo>
                    writer.WriteStartElement("NameInfo");

                    writer.WriteAttributeString("NameTypeCd", NameTypeCd[1]);
                    writer.WriteAttributeString("GivenName", Policy.FirstName);
                    writer.WriteAttributeString("Surname", Policy.LastName);
                    writer.WriteAttributeString("CommercialName", fullName);

                    writer.WriteEndElement(); //End NameInfo
                    #endregion

                    #region <Addr>

                    if(string.IsNullOrEmpty(Policy.MailingAddress) && !string.IsNullOrEmpty(Policy.GaragingAddress))
                    {
                        writer.WriteStartElement("Addr");

                        writer.WriteAttributeString("AddrTypeCd", AddrTypeCd[0]);
                        writer.WriteAttributeString("Addr1", Policy.GaragingAddress);
                        writer.WriteAttributeString("City", Policy.GaragingCity);
                        writer.WriteAttributeString("StateProvCd", Policy.GaragingState);
                        writer.WriteAttributeString("PostalCode", Policy.GaragingZip);

                        writer.WriteEndElement(); //End Addr
                    }
                    else
                    {
                        writer.WriteStartElement("Addr");

                        writer.WriteAttributeString("AddrTypeCd", AddrTypeCd[0]);
                        writer.WriteAttributeString("Addr1", Policy.MailingAddress);
                        writer.WriteAttributeString("City", Policy.MailingCity);
                        writer.WriteAttributeString("StateProvCd", Policy.MailingState);
                        writer.WriteAttributeString("PostalCode", Policy.MailingZip);

                        writer.WriteEndElement(); //End Addr
                    }
                    
                    #endregion

                    #region <Addr>

                    if(Policy.MailingAddress != Policy.GaragingAddress &&
                       Policy.MailingCity != Policy.GaragingCity &&
                       Policy.MailingState != Policy.GaragingState &&
                       Policy.MailingZip != Policy.GaragingZip)
                    {
                        writer.WriteStartElement("Addr");

                        writer.WriteAttributeString("AddrTypeCd", AddrTypeCd[1]);
                        writer.WriteAttributeString("Addr1", Policy.GaragingAddress);
                        writer.WriteAttributeString("City", Policy.GaragingCity);
                        writer.WriteAttributeString("StateProvCd", Policy.GaragingState);
                        writer.WriteAttributeString("PostalCode", Policy.GaragingZip);

                        writer.WriteEndElement(); //End Addr
                    }

                    #endregion

                    writer.WriteEndElement(); //End PartyInfo
                    #endregion

                    writer.WriteEndElement(); //End DTOInsured
                    #endregion

                    #region <DTOBasicPolicy>

                    writer.WriteStartElement("DTOBasicPolicy");

                    writer.WriteAttributeString("CarrierGroupCd", XMLStaticValues.DTOApplication_DTOBasicPolicy_CarrierGroupCd);
                    writer.WriteAttributeString("CarrierCd", XMLStaticValues.DTOApplication_DTOBasicPolicy_CarrierCd);
                    writer.WriteAttributeString("ControllingStateCd", XMLStaticValues.DTOApplication_DTOBasicPolicy_ControllingStateCd);
                    writer.WriteAttributeString("TransactionCd", XMLStaticValues.DTOApplication_DTOBasicPolicy_TransactionCd);
                    writer.WriteAttributeString("InceptionDt", Policy.InceptionDate);
                    writer.WriteAttributeString("InceptionTm", "12:01am");
                    writer.WriteAttributeString("EffectiveDt", Policy.EffectiveDate);
                    writer.WriteAttributeString("EffectiveTm", "12:01am");
                    writer.WriteAttributeString("ExpirationDt", Policy.ExpirationDate);
                    writer.WriteAttributeString("RenewalTermCd", term);
                    writer.WriteAttributeString("ProductVersionIdRef", XMLStaticValues.DTOApplication_DTOBasicPolicy_ProductVersionIdRef);
                    writer.WriteAttributeString("SubTypeCd", XMLStaticValues.DTOApplication_DTOBasicPolicy_SubTypeCd);
                    writer.WriteAttributeString("Description", XMLStaticValues.DTOApplication_DTOBasicPolicy_Description);
                    writer.WriteAttributeString("ProviderNumber", Policy.ProducerCode);
                    writer.WriteAttributeString("TreatAsConvRenewal", XMLStaticValues.DTOApplication_DTOBasicPolicy_TreatAsConvRenewal);
                    writer.WriteAttributeString("LegacyPolNumber", Policy.PolicyNumber);
                    writer.WriteAttributeString("LegacyPolIncepDt", Policy.InceptionDate);
                    writer.WriteAttributeString("Source", XMLStaticValues.DTOApplication_DTOBasicPolicy_Source);
                    writer.WriteAttributeString("RenewedFromPolicyNumber", XMLStaticValues.DTOApplication_DTOBasicPolicy_RenewedFromPolicyNumber);
                    writer.WriteAttributeString("PolicyNumber", XMLStaticValues.DTOApplication_DTOBasicPolicy_PolicyNumber);

                    if (!string.IsNullOrEmpty(Policy.PaymentPlan.Token))
                    {
                        writer.WriteAttributeString("Token", Policy.PaymentPlan.Token);
                    }

                    if (!string.IsNullOrEmpty(Policy.PaymentPlan.SweepDate))
                    {
                        writer.WriteAttributeString("PaymentDay", Policy.PaymentPlan.SweepDate);
                        //writer.WriteAttributeString("MethodCd", Policy.PaymentPlan.EftType);
                    } 
                    

                    #region <ElectronicPaymentSource>

                    writer.WriteStartElement("ElectronicPaymentSource");

                    writer.WriteAttributeString("id", XMLStaticValues.DTOBasicPolicy_ElectronicPaymentSource_id); //TODO
                    writer.WriteAttributeString("SourceTypeCd", XMLStaticValues.DTOBasicPolicy_ElectronicPaymentSource_SourceTypeCd); //TODO
                    //writer.WriteAttributeString("MethodCd", XMLStaticValues.DTOBasicPolicy_ElectronicPaymentSource_MethodCd); //TODO

                    if (!string.IsNullOrEmpty(Policy.PaymentPlan.EftType))
                    {
                        writer.WriteAttributeString("MethodCd", Policy.PaymentPlan.EftType);
                    }                    

                    writer.WriteEndElement(); //End ElectronicPaymentSource

                    #endregion

                    writer.WriteEndElement(); //End DTOBasicPolicy

                    #endregion
                    
                    for(int i = 0; i < Policy.Drivers.Count; i++){

                        if(i != 0){

                            Driver driver = Policy.Drivers[i];
                            int driverNumber = i + 1;
                            tokenizerService.DriverName = string.Format("{0} {1}", driver.FirstName, driver.LastName);
                            var licenseResponse = tokenizerService.Detokenize(driver.LicenseNumber);

                            writer.WriteStartElement("PartyInfo");
                            writer.WriteAttributeString("PartyTypeCd", XMLStaticValues.DTOApplication_PartyInfo_PartyTypeCd);
                            writer.WriteAttributeString("Status", XMLStaticValues.DTOApplication_PartyInfo_Status);
                            
                            writer.WriteStartElement("PersonInfo");

                            writer.WriteAttributeString("PersonTypeCd", PersonTypeCd[0]);
                            writer.WriteAttributeString("GenderCd", driver.Gender);
                            writer.WriteAttributeString("BirthDt", driver.BirthDate);
                            writer.WriteAttributeString("MaritalStatusCd", driver.MaritalStatus);
                            writer.WriteAttributeString("OccupationClassCd", driver.Occupation);
                            //writer.WriteAttributeString("Age", Policy.Age); //TODO

                            writer.WriteEndElement(); //End PersonInfo                                                       

                            writer.WriteStartElement("DriverInfo");

                            writer.WriteAttributeString("DriverInfoCd", XMLStaticValues.PartyInfo_DriverInfo_DriverInfoCd);
                            writer.WriteAttributeString("DriverNumber", driverNumber.ToString());
                            writer.WriteAttributeString("DriverStatusCd", driver.DriverStatus == "Active" ? "Principal" : "Excluded");

                            if (licenseResponse.Result)
                                writer.WriteAttributeString("LicenseNumber", licenseResponse.DetokenizedString);
                            else
                                writer.WriteAttributeString("LicenseNumber", string.Empty);

                            writer.WriteAttributeString("LicenseDt", driver.DateFirstLicense);
                            writer.WriteAttributeString("LicensedStateProvCd", driver.LicenseState);
                            writer.WriteAttributeString("RelationshipToInsuredCd", driver.RelationShip);
                            writer.WriteAttributeString("DriverStartDt", driver.DateFirstLicense);
                            writer.WriteAttributeString("MatureDriverInd", !String.IsNullOrEmpty(driver.MatureDriver) ? driver.MatureDriver : "No");
                            writer.WriteAttributeString("Race", XMLStaticValues.PartyInfo_DriverInfo_Race);
                            writer.WriteAttributeString("LicenseType", XMLStaticValues.PartyInfo_DriverInfo_LicenseType);
                            writer.WriteAttributeString("LicenseStatus", XMLStaticValues.PartyInfo_DriverInfo_LicenseStatus);
                            writer.WriteAttributeString("SR22Ind", !String.IsNullOrEmpty(driver.Sr22) ? driver.Sr22 : "No");
                            writer.WriteAttributeString("OriginalLicenseDt", driver.DateFirstLicense);                            

                            writer.WriteEndElement(); //End DriverInfo

                            writer.WriteStartElement("NameInfo");
                            writer.WriteAttributeString("NameTypeCd", NameTypeCd[0]);
                            writer.WriteAttributeString("GivenName", driver.FirstName);
                            writer.WriteAttributeString("Surname", driver.LastName);
                            writer.WriteEndElement(); //EndNameInfo

                            writer.WriteEndElement(); //End PartyInfo                                                       
                        }                        
                    }

                    for (int j = 0; j < Policy.AdditionalInterests.Count; j++)
                    {
                        
                        int interestNumber = j + 1;
                        AdditionalInterest interest = Policy.AdditionalInterests[j];

                        writer.WriteStartElement("DTOAI"); //DTOAI
                        writer.WriteAttributeString("SequenceNumber", interestNumber.ToString());
                        writer.WriteAttributeString("InterestTypeCd", "Lienholder/Loss Payee");
                        writer.WriteAttributeString("InterestName", interest.Name);
                        writer.WriteAttributeString("Status", "Active");

                        writer.WriteStartElement("PartyInfo"); //PartyInfo 
                        writer.WriteAttributeString("PartyTypeCd", "AIParty");

                        writer.WriteStartElement("EmailInfo"); //EmailInfo 
                        writer.WriteAttributeString("EmailTypeCd", "AIEmail");
                        // writer.WriteAttributeString("PreferredInd", "No");
                        writer.WriteEndElement(); //End EmailInfo

                        writer.WriteStartElement("NameInfo"); //NameInfo 
                        writer.WriteAttributeString("IndexName", interest.Name);
                        writer.WriteAttributeString("NameTypeCd", "AIName");
                        writer.WriteEndElement(); //End NameInfo

                        writer.WriteStartElement("Addr"); //Addr Info 
                        writer.WriteAttributeString("AddrTypeCd", "AIMailingAddr");
                        writer.WriteAttributeString("Addr1", interest.Address);
                        writer.WriteAttributeString("City", interest.City);
                        writer.WriteAttributeString("StateProvCd", interest.State);
                        writer.WriteAttributeString("PostalCode", interest.Zip);                        
                        writer.WriteEndElement(); //End Addr

                        writer.WriteEndElement(); //End PartyInfo

                        writer.WriteEndElement(); //End DTOAI
                    }
                    
                    writer.WriteEndElement(); //End DTOApplication
                    #endregion

                    writer.WriteEndElement(); //End DTORoot

                    writer.WriteEndDocument(); //End Document
                }
            }
            catch(Exception e)
            {
                throw e;
            }
        }

        private void StartElementWithAttributes(XmlWriter writer, string Element, Dictionary<string,string> Attributes)
        {
            writer.WriteStartElement(Element);

            foreach(var attribute in Attributes)
                writer.WriteAttributeString(attribute.Key, attribute.Value);

            writer.WriteEndElement();
        }
    }
}
