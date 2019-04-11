using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.IO;

namespace XMLFileSettings 
{
    public class PropsFields
    {
        public string XMLFileName = Environment.CurrentDirectory + @"\steeings.xml";
        public string IP = "";
    }
    public partial class Props // класс работы с настройками
    {
        public PropsFields Fields;
        public Props()
        {
            Fields = new PropsFields();
        }
        public void writteXML()
        {
            XmlSerializer ser = new XmlSerializer(typeof(PropsFields));
            TextWriter writer = new StreamWriter(Fields.XMLFileName);
            ser.Serialize(writer,Fields);
            writer.Close();
        }
        public void  readerXML()
        {
            if(File.Exists(Fields.XMLFileName))
            {
                XmlSerializer ser = new XmlSerializer(typeof(PropsFields));
                TextReader reader = new StreamReader(Fields.XMLFileName);
                Fields = ser.Deserialize(reader) as PropsFields;
            }
            else
            {
                File.Create(Fields.XMLFileName);
                writteXML();
            }

        }
    }

}
