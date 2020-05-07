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

                    writer.WriteStartElement("QuestionReply");
                    writer.WriteAttributeString("Name", "UWQuestionFraud");
                    writer.WriteAttributeString("Value", "NO");
                    writer.WriteAttributeString("VisibleInd", "YES");
                    writer.WriteEndElement();

                    writer.WriteStartElement("QuestionReply");
                    writer.WriteAttributeString("Name", "UWQuestionVehicleSharing");
                    writer.WriteAttributeString("Value", "NO");
                    writer.WriteAttributeString("VisibleInd", "YES");
                    writer.WriteEndElement();

                    writer.WriteStartElement("QuestionReply");
                    writer.WriteAttributeString("Name", "UWPublicDelivery");
                    writer.WriteAttributeString("Value", "NO");
                    writer.WriteAttributeString("VisibleInd", "YES");
                    writer.WriteEndElement();

                    writer.WriteStartElement("QuestionReply");
                    writer.WriteAttributeString("Name", "UWQuestionOtherDelivery");
                    writer.WriteAttributeString("Value", "NO");
                    writer.WriteAttributeString("VisibleInd", "YES");
                    writer.WriteEndElement();

                    writer.WriteStartElement("QuestionReply");
                    writer.WriteAttributeString("Name", "UWQuestionGaragingLocation");
                    writer.WriteAttributeString("Value", "NO");
                    writer.WriteAttributeString("VisibleInd", "YES");
                    writer.WriteEndElement();

                    writer.WriteStartElement("QuestionReply");
                    writer.WriteAttributeString("Name", "UWQuestionOtherVehicles");
                    writer.WriteAttributeString("Value", "NO");
                    writer.WriteAttributeString("VisibleInd", "YES");
                    writer.WriteEndElement();

                    writer.WriteStartElement("QuestionReply");
                    writer.WriteAttributeString("Name", "UWQuestionHousehold");
                    writer.WriteAttributeString("Value", "NO");
                    writer.WriteAttributeString("VisibleInd", "YES");
                    writer.WriteEndElement();

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
    }
}
