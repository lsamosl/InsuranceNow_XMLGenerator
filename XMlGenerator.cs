using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace InsuranceNow_XMLGenerator
{
    public class XMlGenerator
    {
        public XMlGenerator(string path)
        {
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

                    #region Header

                    writer.WriteStartElement("DTOApplication");
                    writer.WriteAttributeString("Version", "3.10");
                    writer.WriteAttributeString("Status", "In Process");
                    writer.WriteAttributeString("TypeCd", "Application");
                    writer.WriteAttributeString("Description", "New Application");
                    writer.WriteAttributeString("ReadyToRateInd", "Yes");

                    #endregion

                    #region QuestionReplies

                    writer.WriteStartElement("QuestionReplies");
                    writer.WriteAttributeString("QuestionSourceMDA", "UWProduct::product-master::waic-CA-PersonalAuto-v01-00-01:://ProductSetup[@id='ProductSetup']");

                    StartElementWithAttributes(writer, "QuestionReply", 
                           new Dictionary<string, string> { 
                               { "Name", "UWQuestionFraud" } ,
                               { "Value", "NO" } ,
                               { "VisibleInd", "YES" } ,
                           });

                    StartElementWithAttributes(writer, "QuestionReply",
                           new Dictionary<string, string> {
                               { "Name", "UWQuestionVehicleSharing" } ,
                               { "Value", "NO" } ,
                               { "VisibleInd", "YES" } ,
                           });

                    StartElementWithAttributes(writer, "QuestionReply",
                           new Dictionary<string, string> {
                               { "Name", "UWPublicDelivery" } ,
                               { "Value", "NO" } ,
                               { "VisibleInd", "YES" } ,
                           });

                    StartElementWithAttributes(writer, "QuestionReply",
                           new Dictionary<string, string> {
                               { "Name", "UWQuestionOtherDelivery" } ,
                               { "Value", "NO" } ,
                               { "VisibleInd", "YES" } ,
                           });

                    StartElementWithAttributes(writer, "QuestionReply",
                           new Dictionary<string, string> {
                               { "Name", "UWQuestionGaragingLocation" } ,
                               { "Value", "NO" } ,
                               { "VisibleInd", "YES" } ,
                           });

                    StartElementWithAttributes(writer, "QuestionReply",
                           new Dictionary<string, string> {
                               { "Name", "UWQuestionOtherVehicles" } ,
                               { "Value", "NO" } ,
                               { "VisibleInd", "YES" } ,
                           });

                    StartElementWithAttributes(writer, "QuestionReply",
                           new Dictionary<string, string> {
                               { "Name", "UWQuestionHousehold" } ,
                               { "Value", "NO" } ,
                               { "VisibleInd", "YES" } ,
                           });

                    writer.WriteEndElement();

                    #endregion

                    writer.WriteEndElement();
                    writer.WriteEndElement();

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
