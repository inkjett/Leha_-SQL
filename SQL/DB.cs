using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FirebirdSql.Data.FirebirdClient;

namespace SQL
{   
    class DB
    {
        public static FbDataReader ReadData(string queryString)// метод вычитывания данных из бд 
        {
            string connectingString = "character set = WIN1251; initial catalog = " + SQL.Form1.IP + ":" + @"" + SQL.Form1.pathToDB + "; user id = " + SQL.Form1.User + "; password = " + SQL.Form1.Password + "; ";
            try
            {
                SQL.Form1.fb = new FbConnection(connectingString); // записываем строку соединения
                SQL.Form1.fb.Open(); // подключаемся к БД
            }
            catch (Exception e)
            {
                SQL.MessageHelper.GetInstance().SetMessage(e.Message);
            }
            FbTransaction fbt = SQL.Form1.fb.BeginTransaction(); //  начинаем транзакцию данных из БД
            FbCommand SelectSQL = new FbCommand(queryString, SQL.Form1.fb); //запрос
            SelectSQL.Transaction = fbt;
            FbDataReader reader = SelectSQL.ExecuteReader();
            //SelectSQL.Dispose(); //в документации написано, что ОЧЕНЬ рекомендуется убивать объекты этого типа, если они больше не нужны
            return reader;
        }

        public static void CloseFBConnection() // закрыть соединение с БД
        {
            SQL.Form1.fb.Dispose();
            SQL.Form1.fb.Close();
        }
    }
}
