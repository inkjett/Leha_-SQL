using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQL
{
    class DataProcessing
    {


        public void method_arr_of_users(ref List<List<string>> arr_out)//метод формирования массива пользователей
        {
            try
            {
                if (fb.State == ConnectionState.Open)
                {
                    string queryString = "SELECT people.lname||' '||people.fname||' '||people.sname, people.peopleid,cards.cardnum, people.depid FROM cards INNER JOIN people ON(people.peopleid = CARDS.peopleid) where (people.depid != 29) AND (people.depid =" + DEPID + ")";
                    /*
             FbTransaction fbt = fb.BeginTransaction();

             FbCommand SelectSQL = new FbCommand("SELECT people.lname||' '||people.fname||' '||people.sname, people.peopleid,cards.cardnum, people.depid FROM cards INNER JOIN people ON(people.peopleid = CARDS.peopleid) where (people.depid != 29) AND (people.depid =" + DEPID+")", fb); //задаем запрос на выборку исключается ид группы 29 и 40
             SelectSQL.Transaction = fbt;
             FbDataReader reader = SelectSQL.ExecuteReader();*/

                    FbDataReader reader = SQL.DB.ReadData(queryString);
                    List<string> row = new List<string>();
                    Int32 temp = reader.FieldCount;
                    arr_out = new List<List<string>>();

                    try
                    {
                        int i = 0, j = 0;
                        while (reader.Read()) //пока не прочли все данные выполняем... //select_result = select_result + reader.GetInt32(0 ).ToString() + ", " + reader.GetString(1) + "\n";
                        {
                            row = new List<string>();
                            arr_out.Add(row);
                            arr_out[i].Add("");
                            arr_out[i].Add("");
                            arr_out[i].Add("");
                            arr_out[i][j] = reader.GetString(0).ToString();
                            arr_out[i][j + 1] = reader.GetString(1).ToString();
                            arr_out[i][j + 2] = reader.GetString(2).ToString();
                            i++;
                        }
                    }
                    finally
                    {
                        reader.Close();//всегда необходимо вызывать метод Close(), когда чтение данных завершено
                        data_is_read = true;
                        //fb.Close(); //закрываем соединение, т.к. оно нам больше не нужно
                    }
                    //SelectSQL.Dispose();//в документации написано, что ОЧЕНЬ рекомендуется убивать объекты этого типа, если они больше не нужны
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Сообщение", MessageBoxButtons.OK);
            }

        }




    }
}
