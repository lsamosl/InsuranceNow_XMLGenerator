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

                using (XmlWriter writer = XmlWriter.Create(Path, settings))
                {
                    writer.WriteStartDocument();

                    writer.WriteStartElement("DTORoot");

                    writer.WriteStartElement("DTOApplication");
                    writer.WriteAttributeString("Version", XMLStaticValues.DTORoot_DTOApplication_Version);
                    writer.WriteAttributeString("Status", XMLStaticValues.DTORoot_DTOApplication_Status);
                    writer.WriteAttributeString("TypeCd", XMLStaticValues.DTORoot_DTOApplication_TypeCd);
                    writer.WriteAttributeString("Description", XMLStaticValues.DTORoot_DTOApplication_Description);
                    writer.WriteAttributeString("ReadyToRateInd", XMLStaticValues.DTORoot_DTOApplication_ReadyToRateInd);

                    writer.WriteStartElement("QuestionReplies");
                    writer.WriteAttributeString("QuestionSourceMDA", XMLStaticValues.DTOApplication_QuestionReplies_QuestionSourceMDA);

                    string[] QuestionReply_Names = XMLStaticValues.QuestionReplies_QuestionReply_Name.Split('|');

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

                    writer.WriteStartElement("DTOLine");

                    string biCover = Policy.General.BiCover;
                    string[] biCoverArray = biCover.Split('/');

                    writer.WriteAttributeString("StatusCd", XMLStaticValues.DTOApplication_DTOLine_StatusCd);
                    writer.WriteAttributeString("LineCd", XMLStaticValues.DTOApplication_DTOLine_LineCd);
                    writer.WriteAttributeString("RatingService", XMLStaticValues.DTOApplication_DTOLine_RatingService);
                    writer.WriteAttributeString("MedPayLimit", !String.IsNullOrEmpty(Policy.General.MedCover) ? Policy.General.MedCover : "None");
                    writer.WriteAttributeString("BILimit", biCoverArray.Length == 2 ? biCoverArray[0] + "000" + "/" + biCoverArray[1] + "000" : "None");
                    writer.WriteAttributeString("PDLimit", !String.IsNullOrEmpty(Policy.General.PdCover) ? Policy.General.PdCover : "None");
                    writer.WriteAttributeString("UMBILimit", !String.IsNullOrEmpty(Policy.General.Umbi) ? Policy.General.Umbi : "None");
                    writer.WriteAttributeString("UMPDWCDInd", !String.IsNullOrEmpty(Policy.General.UmpdCdw) ? Policy.General.UmpdCdw : "No");
                    writer.WriteAttributeString("LimitedMexicoInd", !String.IsNullOrEmpty(Policy.General.LimMex) ? Policy.General.LimMex : "No");
                    writer.WriteAttributeString("RoadAssistInd", !String.IsNullOrEmpty(Policy.General.RoadAssis) ? Policy.General.RoadAssis : "No");
                    writer.WriteAttributeString("PolicyType", XMLStaticValues.DTOApplication_DTOLine_PolicyType);
                    writer.WriteAttributeString("TotalNumVehicles", Policy.Vehicles.Count.ToString());
                    writer.WriteAttributeString("TotalNumDrivers", Policy.Drivers.Count.ToString());

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

                    foreach(Vehicle v in Policy.Vehicles)
                    {

                    }

                    writer.WriteEndElement(); //End DTORisk

                    writer.WriteEndElement(); //End DTOLine

                    writer.WriteEndElement(); //End DTOApplication

                    writer.WriteEndElement(); //End DTORoot

                    writer.WriteEndDocument(); 
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
