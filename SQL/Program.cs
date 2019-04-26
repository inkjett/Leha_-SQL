using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using System.IO;

namespace SQL
{
    static class Program
    {
        public static Form1 f1;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        //запись данных в хмл

        public class PropsFields
        {
            public string XMLFileName = Environment.CurrentDirectory + @"\steeings.xml";
            public string IP = "192.168.0.1";
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
                ser.Serialize(writer, Fields);
                writer.Close();
            }
            public void readerXML()
            {
                if (File.Exists(Fields.XMLFileName))
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
}
