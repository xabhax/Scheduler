using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Scheduler
{
    class XMLFunctions
    {
        public List<string> TeamOne;
        public List<string> TeamTwo;
        public List<string> TeamThree;    

        public static string XMLFile = System.Environment.CurrentDirectory + "\\schedule.dat";

        public static void Read()
        {

        }

        public static void Write()
        {
            XmlWriterSettings st = new XmlWriterSettings();
            st.Indent = true;

            st.OmitXmlDeclaration = true;
            st.Encoding = Encoding.ASCII;
            string path = "config.xml";

            using (XmlWriter writer = XmlWriter.Create(path, st))
            {
                writer.WriteStartElement("Schedule");
                writer.WriteStartElement("Team1");
                writer.WriteValue("Jae Park");
                writer.WriteValue("Fischer");
                writer.WriteValue("David");
                writer.WriteValue("Powers");
                writer.WriteStartElement("Team2");
                writer.WriteValue("Rich");
                writer.WriteValue("Thompson");
                writer.WriteValue("Matt");
                writer.WriteValue("Whiskers");
                writer.WriteStartElement("Team3");
                writer.WriteValue("Marco");
                writer.WriteValue("Jay Jay");
                writer.WriteValue("Kevin");
                writer.WriteValue("Brian");
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
        }
    }
}
