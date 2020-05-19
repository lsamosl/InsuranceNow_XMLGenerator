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
