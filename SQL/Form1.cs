using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FirebirdSql.Data.FirebirdClient;
using System.Collections;
using System.IO;

namespace SQL
{
    public partial class Form1 : Form
    {
        FbConnection fb;
        List<List<string>> arr_user;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
               

        private void button3_Click(object sender, EventArgs e)
        {
            FbConnectionStringBuilder fb_connect = new FbConnectionStringBuilder();
            fb_connect.Charset = "WIN1251"; // кодировка
            fb_connect.UserID = textBox2.Text; // Логин
            fb_connect.Password = textBox3.Text; // Пароль
            fb_connect.Database = textBox1.Text; // путь до БД
            fb_connect.ServerType = 0; //  хз что такое  ----   указываем тип сервера (0 - "полноценный Firebird" (classic или super server), 1 - встроенный (embedded))
            fb = new FbConnection(fb_connect.ToString()); // открываем подключение, вставляя строку подключения 
            fb.Open();
            //FbDatabaseInfo fb_info = new FbDatabaseInfo(fb);
            //MessageBox.Show("Info: "+ fb_info.ServerClass+"\nVer: "+fb_info.ServerVersion);
            if(fb.State == ConnectionState.Open)
            {
                label5.Text = "Подключено";
                label5.ForeColor = Color.Green;
            }
            else
            {
                label5.Text = "Что-то пошло не так..";
                label5.ForeColor = Color.Red;
            }
        }

        private void button2_Click(object sender, EventArgs e)//select
        {
            if (fb.State == ConnectionState.Open)
            {
                int i = 0, j = 0;

                FbTransaction fbt = fb.BeginTransaction();
                FbCommand SelectSQL = new FbCommand("SELECT people.lname||' '||people.fname||' '||people.sname, people.peopleid,cards.cardnum FROM cards INNER JOIN people ON(people.peopleid = CARDS.peopleid)", fb); //задаем запрос на выборку
                SelectSQL.Transaction = fbt;
                FbDataReader reader = SelectSQL.ExecuteReader();
                string select_result = "";

                List<string> row = new List<string>();
                Int32 temp = reader.FieldCount;
                arr_user = new List<List<string>>();
                
                try
                {
                    while (reader.Read()) //пока не прочли все данные выполняем... //select_result = select_result + reader.GetInt32(0 ).ToString() + ", " + reader.GetString(1) + "\n";
                    {
                        row = new List<string>();
                        arr_user.Add(row);
                        arr_user[i].Add("");
                        arr_user[i].Add("");
                        arr_user[i].Add("");
                        arr_user [i][j]= reader.GetString(0).ToString();
                        arr_user[i][j+1] = reader.GetString(1).ToString();
                        arr_user[i][j + 2] = reader.GetString(2).ToString();
                        i++;
                    }
                }
                finally
                {
                    //всегда необходимо вызывать метод Close(), когда чтение данных завершено
                    reader.Close();
                    fb.Close(); //закрываем соединение, т.к. оно нам больше не нужно
                }
                SelectSQL.Dispose();//в документации написано, что ОЧЕНЬ рекомендуется убивать объекты этого типа, если они больше не нужны


                dataGridView1.RowCount = arr_user.Count;
                dataGridView1.ColumnCount = 3;
                for (int ii=0;ii<arr_user.Count;ii++)
                {
                    for (int jj=0;jj<3;jj++)
                    {
                        dataGridView1.Rows[ii].Cells[jj].Value = String.Format("{0}",arr_user[ii][jj]);
                    }

                }




            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

        }
    }
}
