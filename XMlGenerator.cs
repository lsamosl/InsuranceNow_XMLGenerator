using InsuranceNow_XMLGenerator.Models;
using InsuranceNow_XMLGenerator.Resources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace InsuranceNow_XMLGenerator
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
                string BILimit = GetFormatLimits(Policy.General.BiCover);
                //string UMBILimit = GetFormatLimits(Policy.General.Umbi);

                using (XmlWriter writer = XmlWriter.Create(Path, settings))
                {
                    writer.WriteStartDocument();

                    writer.WriteStartElement("DTORoot");

                    #region <DTOApplication>
                    writer.WriteStartElement("DTOApplication");

                    writer.WriteAttributeString("Version", XMLStaticValues.DTORoot_DTOApplication_Version);
                    writer.WriteAttributeString("Status", XMLStaticValues.DTORoot_DTOApplication_Status);
                    writer.WriteAttributeString("TypeCd", XMLStaticValues.DTORoot_DTOApplication_TypeCd);
                    writer.WriteAttributeString("Description", XMLStaticValues.DTORoot_DTOApplication_Description);
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
                    writer.WriteAttributeString("MedPayLimit", !String.IsNullOrEmpty(Policy.General.MedCover) ? Policy.General.MedCover : "None");
                    writer.WriteAttributeString("BILimit", BILimit);
                    writer.WriteAttributeString("PDLimit", !String.IsNullOrEmpty(Policy.General.PdCover) ? Policy.General.PdCover : "None");
                    writer.WriteAttributeString("UMBILimit", !String.IsNullOrEmpty(Policy.General.Umbi) ? Policy.General.Umbi : "None");
                    writer.WriteAttributeString("UMPDWCDInd", !String.IsNullOrEmpty(Policy.General.UmpdCdw) ? Policy.General.UmpdCdw : "No");
                    writer.WriteAttributeString("LimitedMexicoInd", !String.IsNullOrEmpty(Policy.General.LimMex) ? Policy.General.LimMex : "No");
                    writer.WriteAttributeString("RoadAssistInd", !String.IsNullOrEmpty(Policy.General.RoadAssis) ? Policy.General.RoadAssis : "No");
                    writer.WriteAttributeString("PolicyType", XMLStaticValues.DTOApplication_DTOLine_PolicyType);
                    writer.WriteAttributeString("TotalNumVehicles", Policy.Vehicles.Count.ToString());
                    writer.WriteAttributeString("TotalNumDrivers", Policy.Drivers.Count.ToString());

                    #region <DTORisk>
                    writer.WriteStartElement("DTORisk");

                    writer.WriteAttributeString("TypeCd", XMLStaticValues.DTOLine_DTORisk_TypeCd);
                    writer.WriteAttributeString("Status", XMLStaticValues.DTOLine_DTORisk_Status);
                    writer.WriteAttributeString("Description", string.Empty);
                    writer.WriteAttributeString("RiskBeanName", string.Empty);
                    writer.WriteAttributeString("RiskAddDt", string.Empty);
                    writer.WriteAttributeString("RiskAddPolicyVersion", XMLStaticValues.DTOLine_DTORisk_RiskAddPolicyVersion);
                    writer.WriteAttributeString("RiskAddTransactionNo", XMLStaticValues.DTOLine_DTORisk_RiskAddTransactionNo);

                    writer.WriteStartElement("DTOCoverage");

                    writer.WriteAttributeString("Status", XMLStaticValues.DTORisk1_DTOCoverage1_Status);
                    writer.WriteAttributeString("CoverageCd", XMLStaticValues.DTORisk1_DTOCoverage1_CoverageCd);
                    writer.WriteAttributeString("Description", XMLStaticValues.DTORisk1_DTOCoverage1_Description);

                    string[] limits = XMLStaticValues.DTOCoverage1_DTOLimit_LimitCd.Split('|');
                    string[] limits_values = XMLStaticValues.DTOCoverage1_DTOLimit_TypeCd.Split('|');

                    for(int i = 0; i< limits.Length; i++){

                        writer.WriteStartElement("DTOLimit");

                        writer.WriteAttributeString("LimitCd", limits[i]);
                        writer.WriteAttributeString("TypeCd", limits_values[i]);
                        
                        if(biCoverArray.Length == limits.Length)
                            writer.WriteAttributeString("Value", biCoverArray[i] + "000");  
                        else
                            writer.WriteAttributeString("Value", string.Empty);

                        writer.WriteEndElement(); //End DTOLimit
                    }

                    writer.WriteEndElement(); //End DTOCoverage

                    #region <DTOVehicle>
                    foreach (Vehicle v in Policy.Vehicles)
                    {
                        writer.WriteStartElement("DTOVehicle");

                        writer.WriteAttributeString("VehNumber", countVehicles.ToString());
                        writer.WriteAttributeString("Status", XMLStaticValues.DTORisk_DTOVehicle_Status);
                        writer.WriteAttributeString("VehIdentificationNumber", v.VIN);
                        writer.WriteAttributeString("ValidVinInd", XMLStaticValues.DTORisk_DTOVehicle_ValidVinInd);
                        writer.WriteAttributeString("VehUseCd", XMLStaticValues.DTORisk_DTOVehicle_VehUseCd);
                        writer.WriteAttributeString("CollisionDed", !string.IsNullOrEmpty(v.CollDeduct) ? v.CollDeduct : "No Coverage");
                        writer.WriteAttributeString("ComprehensiveDed", !string.IsNullOrEmpty(v.CompDeduct) ? v.CompDeduct : "No Coverage");
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
                    }
                    #endregion

                    writer.WriteStartElement("DTOCoverage");

                    writer.WriteAttributeString("Status", XMLStaticValues.DTORisk2_DTOCoverage2_Status);
                    writer.WriteAttributeString("CoverageCd", XMLStaticValues.DTORisk2_DTOCoverage2_CoverageCd);
                    writer.WriteAttributeString("Description", XMLStaticValues.DTORisk2_DTOCoverage2_Description);

                    writer.WriteStartElement("DTOLimit");

                    writer.WriteAttributeString("LimitCd", XMLStaticValues.DTOCoverage2_DTOLimit_LimitCd);
                    writer.WriteAttributeString("TypeCd", XMLStaticValues.DTOCoverage2_DTOLimit_TypeCd);
                    writer.WriteAttributeString("Value", !String.IsNullOrEmpty(Policy.General.PdCover) ? Policy.General.PdCover : "None");

                    writer.WriteEndElement(); //End DTOLimit

                    writer.WriteEndElement(); //End DTOCoverage

                    writer.WriteEndElement(); //End DTORisk
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
                    writer.WriteAttributeString("Age", string.Empty); //TODO

                    writer.WriteEndElement(); //End PersonInfo
                    #endregion

                    #region <DriverInfo>
                    
                    foreach(Driver d in Policy.Drivers)
                    {
                        writer.WriteStartElement("DriverInfo");

                        writer.WriteAttributeString("DriverInfoCd", XMLStaticValues.PartyInfo_DriverInfo_DriverInfoCd); 
                        writer.WriteAttributeString("DriverNumber", countDrivers.ToString());
                        writer.WriteAttributeString("DriverStatusCd", d.DriverStatus == "Active" ? "Principal"  : "Excluded"); 
                        writer.WriteAttributeString("LicenseNumber", d.LicenseNumber); 
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
                    }
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
                    writer.WriteAttributeString("TransactionEffectiveDt", string.Empty); //TODO
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

                    writer.WriteStartElement("Addr");

                    writer.WriteAttributeString("AddrTypeCd", AddrTypeCd[0]);
                    writer.WriteAttributeString("Addr1", Policy.MailingAddress);
                    writer.WriteAttributeString("City", Policy.MailingCity);
                    writer.WriteAttributeString("StateProvCd", Policy.MailingState);
                    writer.WriteAttributeString("PostalCode", Policy.MailingZip);

                    writer.WriteEndElement(); //End Addr
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
                    writer.WriteAttributeString("ProviderNumber", string.Format("{0}-{1}", Policy.ProducerCode, "001"));
                    writer.WriteAttributeString("TreatAsConvRenewal", XMLStaticValues.DTOApplication_DTOBasicPolicy_TreatAsConvRenewal);
                    writer.WriteAttributeString("LegacyPolNumber", Policy.PolicyNumber);
                    writer.WriteAttributeString("LegacyPolIncepDt", Policy.InceptionDate);
                    writer.WriteAttributeString("Source", XMLStaticValues.DTOApplication_DTOBasicPolicy_Source);
                    writer.WriteAttributeString("RenewedFromPolicyNumber", XMLStaticValues.DTOApplication_DTOBasicPolicy_RenewedFromPolicyNumber);
                    writer.WriteAttributeString("PolicyNumber", XMLStaticValues.DTOApplication_DTOBasicPolicy_PolicyNumber);

                    writer.WriteEndElement(); //End DTOBasicPolicy

                    #region <ElectronicPaymentSource>

                    writer.WriteStartElement("ElectronicPaymentSource");

                    writer.WriteAttributeString("id", XMLStaticValues.DTOBasicPolicy_ElectronicPaymentSource_id); //TODO
                    writer.WriteAttributeString("SourceTypeCd", XMLStaticValues.DTOBasicPolicy_ElectronicPaymentSource_SourceTypeCd); //TODO
                    writer.WriteAttributeString("MethodCd", XMLStaticValues.DTOBasicPolicy_ElectronicPaymentSource_MethodCd); //TODO

                    writer.WriteEndElement(); //End ElectronicPaymentSource

                    #endregion

                    #endregion

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

        private string GetFormatLimits(string limits)
        {
            string limit = "None";

            if (limits.ToLower() == "no" || limits.ToLower() == "yes")
                return limits;

            string[] limitsArray = limits.Split('/');
            if (limitsArray.Length == 2)
                limit = string.Format("{0}/{1}", limitsArray[0] + "000", limitsArray[1] + "000");

            return limit;
        }
    }
}
