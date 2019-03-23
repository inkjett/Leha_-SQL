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
using Excel = Microsoft.Office.Interop.Excel;

namespace SQL
{
    public partial class Form1 : Form
    {

        //описание переменных
        FbConnection fb;
        public Excel.Application excelapp;
        public Excel.Worksheet excelworksheet;
        public Excel.Workbooks excelappworkbooks;
        public Excel.Workbook excelappworkbook;
        public Excel.Sheets excelsheets;
        public Excel.Range excelcells;
        List<List<string>> arr_user;
        List<List<string>> arr_events;
        List<List<string>> arr_of_work;
        string date_to_request = "0";
        public string path_to_DB = "C:\\123.fdb";
        public string User = "SYSDBA";
        public string Pass = "masterkey";
        DateTime now = DateTime.Now;
       
        
        //------
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }


        //методы
        public void method_connect_to_fb(TextBox Text_Path,TextBox Text_user,TextBox Text_pass, ref Label Label_out)// метод подключения к БД
        {
            FbConnectionStringBuilder fb_connect = new FbConnectionStringBuilder();
            fb_connect.Charset = "WIN1251"; // кодировка
            fb_connect.UserID = Text_user.Text; // Логин
            fb_connect.Password = Text_pass.Text; // Пароль
            fb_connect.Database = Text_Path.Text; // путь до БД
            fb_connect.ServerType = 0; //  хз что такое  ----   указываем тип сервера (0 - "полноценный Firebird" (classic или super server), 1 - встроенный (embedded))
            fb = new FbConnection(fb_connect.ToString()); // открываем подключение, вставляя строку подключения 
            fb.Open();
            //FbDatabaseInfo fb_info = new FbDatabaseInfo(fb);
            //MessageBox.Show("Info: "+ fb_info.ServerClass+"\nVer: "+fb_info.ServerVersion);
            if (fb.State == ConnectionState.Open)
            {
                Label_out.Text = "Подключено";
                Label_out.ForeColor = Color.Green;
            }
            else
            {
                Label_out.Text = "Что-то пошло не так..";
                Label_out.ForeColor = Color.Red;
            }
        }

        public void method_arr_of_users(ref List<List<string>> arr_out)//метод формирования массива пользователей
        {
            
            if (fb.State == ConnectionState.Open)
            {
                int i = 0, j = 0;

                FbTransaction fbt = fb.BeginTransaction();
                FbCommand SelectSQL = new FbCommand("SELECT people.lname||' '||people.fname||' '||people.sname, people.peopleid,cards.cardnum, people.depid FROM cards INNER JOIN people ON(people.peopleid = CARDS.peopleid) where people.depid != 29", fb); //задаем запрос на выборку исключается ид группы 29 и 40
                SelectSQL.Transaction = fbt;
                FbDataReader reader = SelectSQL.ExecuteReader();
                
                List<string> row = new List<string>();
                Int32 temp = reader.FieldCount;
                arr_out = new List<List<string>>();

                try
                {
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
                    //всегда необходимо вызывать метод Close(), когда чтение данных завершено
                    reader.Close();
                    fb.Close(); //закрываем соединение, т.к. оно нам больше не нужно
                }
                SelectSQL.Dispose();//в документации написано, что ОЧЕНЬ рекомендуется убивать объекты этого типа, если они больше не нужны
            }
        }

        
        public void method_arr_of_events(string date_to_request_in, ref List<List<string>> arr_out)//метод формирования массива о отработанном времени 
        {
            bool start_bool = true;
            if (fb.State == ConnectionState.Open)
            {
                int i = 0, j = 0;

                FbTransaction fbt = fb.BeginTransaction();
                FbCommand SelectSQL = new FbCommand("SELECT DISTINCT events.eventsdate, events.cardnum,events.readerid FROM events WHERE events.eventsdate >= '" + date_to_request_in + " 00:00:00' AND events.eventsdate <= '" + date_to_request_in + " 23:59:59'", fb); //задаем запрос на выборку
                SelectSQL.Transaction = fbt;
                FbDataReader reader = SelectSQL.ExecuteReader();

                List<string> row = new List<string>();
                Int32 temp = reader.FieldCount;
                arr_out = new List<List<string>>();

                try
                {
                    while (reader.Read()) //пока не прочли все данные выполняем...
                    {
                        row = new List<string>();
                        arr_out.Add(row);
                        arr_out[i].Add("");
                        arr_out[i].Add("");
                        arr_out[i].Add("");
                        arr_out[i][j] = reader.GetString(0).ToString();//добавиои время
                        arr_out[i][j + 1] = reader.GetString(1).ToString();// добавили id ключа
                        arr_out[i][j + 2] = reader.GetString(2).ToString();//добавили id эвента
                        for (int ii = 0; (ii < arr_out.Count - 1) && start_bool == false; ii++)// запуск цикла проверки на существующие запиви
                        {
                            if ((arr_out[ii][1] == arr_out[i][1])&& (arr_out[ii][2] == arr_out[i][2]))//проверка на уже существующее ID 
                            {
                                if (Convert.ToInt32(arr_out[i][2]) == 3)// проверка на сощуствующий эвент 3-вход 13-выход
                                {
                                    if (DateTime.Parse(arr_out[ii][0]) > DateTime.Parse(arr_out[i][0])) // проверка на время
                                    {
                                        arr_out[ii][0] = arr_out[i][0];
                                    }
                                }
                                if (Convert.ToInt32(arr_out[i][2]) == 13)// провоека на сощуствующий эвент 3-вход 13-выход
                                {
                                    if (DateTime.Parse(arr_out[ii][0]) < DateTime.Parse(arr_out[i][0])) // проверка на время
                                    {
                                        arr_out[ii][0] = arr_out[i][0];
                                    }
                                }
                                arr_out.Remove(row);
                                i--;
                            }
                        }
                        i++;
                        start_bool = false;
                    }
                }
                finally
                {
                    //всегда необходимо вызывать метод Close(), когда чтение данных завершено
                    reader.Close();
                    fb.Close(); //закрываем соединение, т.к. оно нам больше не нужно
                }
                SelectSQL.Dispose();//в документации написано, что ОЧЕНЬ рекомендуется убивать объекты этого типа, если они больше не нужны
            }                                                  
        }
               
        public void method_arr_to_grid(List<List<string>> arr_in, ref DataGridView Grid_out)//метод вывывода массива в датагрид
        {
            if (arr_in.Count != 0)
            {
                Grid_out.RowCount = arr_in.Count;
                Grid_out.ColumnCount = 3;
                dataGridView3.Columns[0].Width = 180;
                dataGridView3.Columns[1].Width = 120;
                dataGridView3.Columns[2].Width = 120;
                for (int ii = 0; ii < arr_in.Count; ii++)
                {
                    for (int jj = 0; jj < 3; jj++)
                    {
                        Grid_out.Rows[ii].Cells[jj].Value = String.Format("{0}", arr_in[ii][jj]);
                    }

                }
            }
            else
            {
                MessageBox.Show("Данных за данный период нет.");
            }
        }


        public void method_of_end_arr(List<List<string>> arr_events_in, List<List<string>> arr_user_in, ref List<List<string>> arr_out)//метод формирования массива по отработанному времени
        {
            List<string> row = new List<string>();
            arr_out = new List<List<string>>();

            for (int i = 0; i < arr_user_in.Count; i++)
            {
                row = new List<string>();
                arr_out.Add(row);
                arr_out[i].Add("");
                arr_out[i].Add("");
                arr_out[i].Add("");
                arr_out[i][0] = arr_user_in[i][0];


                for (int ii = 0; ii < arr_events_in.Count; ii++)
                {
                    if ((arr_user_in[i][2] == arr_events_in[ii][1]) && (Convert.ToInt32(arr_events_in[ii][2]) == 3))
                    {
                        arr_out[i][1] = arr_events_in[ii][0];
                    }
                    if ((arr_user_in[i][2] == arr_events_in[ii][1]) && (Convert.ToInt32(arr_events_in[ii][2]) == 13))
                    {
                        arr_out[i][2] = arr_events_in[ii][0];
                    }
                }
            }



        }        

        //-------------


        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            date_to_request = dateTimePicker1.Value.ToShortDateString();
        }


        private void button7_Click(object sender, EventArgs e)
        {


            try
                {
                    method_connect_to_fb(textBox1, textBox2, textBox3, ref label5);
                }
                catch
                {
                    MessageBox.Show("Проверьте настройки подключения", "Сообщение", MessageBoxButtons.OK);

                }

            try
            {
            if (date_to_request == "0")
               {
                    date_to_request = now.ToString("dd.MM.yyyy");
               }
                method_arr_of_users(ref arr_user);
            method_connect_to_fb(textBox1, textBox2, textBox3, ref label5);
            method_arr_of_events(date_to_request, ref arr_events);
            method_connect_to_fb(textBox1, textBox2, textBox3, ref label5);
            method_of_end_arr(arr_events, arr_user, ref arr_of_work);
            method_arr_to_grid(arr_of_work, ref dataGridView3);
            }
            catch
            {
                MessageBox.Show("Что-то пошло не так...", "Сообщение", MessageBoxButtons.OK);
            }     

        }


        private void button4_Click(object sender, EventArgs e)
        {
            Close(); 
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int start_arr_in_excel = 3;
            excelapp = new Excel.Application();// создаем новую книгу
            excelapp.Visible = true;
            excelapp.SheetsInNewWorkbook = 1; // указываем количество листов
            excelapp.Workbooks.Add();//добавляем лист 
            excelappworkbooks = excelapp.Workbooks;//определяем книгу(вроде как)
            excelappworkbook = excelappworkbooks[1];//Получаем ссылку на книгу 1 - нумерация от 1
            excelsheets = excelappworkbook.Worksheets;//определяем листы
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);//получаем ссылку на первый лист
            excelworksheet.Activate();//делаем активным первый лист 


            excelcells = excelworksheet.get_Range("A1", "D1");
            excelcells.Select();
            ((Excel.Range)(excelapp.Selection)).Merge(Type.Missing);
            excelcells = (Excel.Range)excelworksheet.Cells[1, "A"];//выбираем ячейку для заполнения
            excelcells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
            excelcells.Value2 = "Отчет о времени проведенном на рабочем месте за " + date_to_request;
            excelcells = (Excel.Range)excelworksheet.Cells[2, "A"];//выбираем ячейку для заполнения
            excelcells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
            excelcells.Value2 = "№";
            excelcells.Cells.ColumnWidth = 3;
            excelcells = (Excel.Range)excelworksheet.Cells[2, "B"];//выбираем ячейку для заполнения
            excelcells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
            excelcells.Value2 = "Фамилия Имя Отчество";
            excelcells.Cells.ColumnWidth = 35;
            excelcells = (Excel.Range)excelworksheet.Cells[2, "C"];//выбираем ячейку для заполнения
            excelcells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
            excelcells.Value2 = "Время прихда на работу";
            excelcells.Cells.ColumnWidth = 21;
            excelcells = (Excel.Range)excelworksheet.Cells[2, "D"];//выбираем ячейку для заполнения
            excelcells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
            excelcells.Value2 = "Время ухода с работы";
            excelcells.Cells.ColumnWidth = 21;



            for (int i = 0; i < arr_of_work.Count; i++)
            {
                excelcells = (Excel.Range)excelworksheet.Cells[i + start_arr_in_excel, "A"];//выбираем ячейку для заполнения
                excelcells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
                excelcells.Value2 = i + 1;
                excelcells = (Excel.Range)excelworksheet.Cells[i + start_arr_in_excel, "B"];//выбираем ячейку для заполнения
                excelcells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
                excelcells.Value2 = arr_of_work[i][0];
                excelcells = (Excel.Range)excelworksheet.Cells[i + start_arr_in_excel, "C"];//выбираем ячейку для заполнения
                excelcells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
                excelcells.Value2 = arr_of_work[i][1];
                excelcells = (Excel.Range)excelworksheet.Cells[i + start_arr_in_excel, "D"];//выбираем ячейку для заполнения
                excelcells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
                excelcells.Value2 = arr_of_work[i][2];
            }



        }
    }
}
