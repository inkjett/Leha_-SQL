using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.IO;

namespace OPTsettings 
{
    public class PropsFields
    {        
        public string IP;
        public string pathToDB;
        public string User;
        public string Password;
    }
    public class Props // класс работы с настройками
    {
        public static string XMLFileName = Environment.CurrentDirectory + @"\steeings.xml";
        
        public static PropsFields Fields;
        public Props()
        {
            Fields = new PropsFields();
        }
        public static void CopyItemsToSer() // копирование полей с настройками из текущих настроек подключения
        {
            Fields.IP = SQL.Form1.IP;
            Fields.pathToDB = SQL.Form1.pathToDB;
            Fields.User = SQL.Form1.User;
            Fields.Password = SQL.Form1.Password;
        }
        public static void CopyItemsToProgramm() // копирование полей с настройками из текущих настроек подключения
        {
            SQL.Form1.IP = Fields.IP;
            SQL.Form1.pathToDB = Fields.pathToDB;
            SQL.Form1.User = Fields.User;
            SQL.Form1.Password = Fields.Password;
        }
        public static void writteXML()
        {
            var props = new Props();
            CopyItemsToSer();
            XmlSerializer ser = new XmlSerializer(typeof(PropsFields));
            TextWriter writer = new StreamWriter(XMLFileName);
            ser.Serialize(writer,Fields);
            writer.Close();
        }
        public void readerXML()
        {
            if (File.Exists(XMLFileName))
            {
                XmlSerializer ser = new XmlSerializer(typeof(PropsFields));
                TextReader reader = new StreamReader(XMLFileName);
                Fields = ser.Deserialize(reader) as PropsFields;
                CopyItemsToProgramm();
                reader.Close();
            }
            else
            {
                SQL.MessageHelper.GetInstance().SetMessage("Файл настроек не найден, исползуются стандарные настройки");
            }

        }
    }

}
