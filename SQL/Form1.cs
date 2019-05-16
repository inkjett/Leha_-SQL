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
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Text.RegularExpressions;
using System.Threading;
using System.Timers;

namespace SQL
{
    public partial class Form1 : Form
    {

        //описание переменных
        FbConnection fb;
        public List<List<string>> arr_user;
        List<List<string>> arr_events;
        List<List<string>> arr_events_per_mounth;
        List<List<string>> arr_of_work;
        public List<List<string>> arr_of_deviation;  
        string date_to_request = "0";
        public bool data_is_read = false;
        public string path_to_DB = "C:\\123.fdb";
        public string User = "SYSDBA";
        public string Pass = "masterkey";
        DateTime now = DateTime.Now;
        TimeSpan infinite = TimeSpan.FromMilliseconds(-1);
        TimeSpan hour = new TimeSpan(1, 0, 0);
        TimeSpan four_hour = new TimeSpan(4, 0, 0);
        public string connecting_path;
                              
        //------
        public Form1()
        {
            InitializeComponent();
            Program.f1 = this;            
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

    
        //методы
        public string method_connection_string (TextBox Text_Path, TextBox Text_user, TextBox Text_pass)
        {
            string path_path = "character set = WIN1251; initial catalog = " + textBox4.Text + ":" + @"" + textBox1.Text + "; user id = " + Text_user.Text + "; password = " + Text_pass.Text + "; ";
            return path_path;
        }



        public void method_connect_to_fb(string path_in, ref Label Label_out)// метод подключения к БД
        {
            try
            {
                //string path_path = "character set = WIN1251; initial catalog = "+textBox4.Text+":" + @""+textBox1.Text+"; user id = " + Text_user.Text + "; password = " + Text_pass.Text + "; ";
                fb = new FbConnection(path_in);
                fb.Open();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Сообщение", MessageBoxButtons.OK);
            }
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
            try
            {
                if (fb.State == ConnectionState.Open)
                {
                    int i = 0, j = 0;

                    FbTransaction fbt = fb.BeginTransaction();
                    FbCommand SelectSQL = new FbCommand("SELECT people.lname||' '||people.fname||' '||people.sname, people.peopleid,cards.cardnum, people.depid FROM cards INNER JOIN people ON(people.peopleid = CARDS.peopleid) where (people.depid != 29) AND (people.depid != 40)", fb); //задаем запрос на выборку исключается ид группы 29 и 40
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
                        fbt.Dispose();
                        data_is_read = true;
                        //fb.Close(); //закрываем соединение, т.к. оно нам больше не нужно
                    }
                    SelectSQL.Dispose();//в документации написано, что ОЧЕНЬ рекомендуется убивать объекты этого типа, если они больше не нужны
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Сообщение", MessageBoxButtons.OK);
            }

        }


        public void method_arr_of_events(string date_to_request_in, ref List<List<string>> arr_out)//метод формирования массива о отработанном времени 
        {
            try
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
                                if ((arr_out[ii][1] == arr_out[i][1]) && (arr_out[ii][2] == arr_out[i][2]))//проверка на уже существующее ID 
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
                        fbt.Dispose();
                        //fb.Close(); //закрываем соединение, т.к. оно нам больше не нужно
                    }
                    SelectSQL.Dispose();//в документации написано, что ОЧЕНЬ рекомендуется убивать объекты этого типа, если они больше не нужны
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Сообщение", MessageBoxButtons.OK);
            }

        }

        public void method_arr_to_grid(List<List<string>> arr_in, ref DataGridView Grid_out)//метод вывывода массива в датагрид
        {
            if (data_is_read == true)
            {
                try
                {
                    if (arr_in.Count != 0)
                    {
                        Grid_out.RowCount = arr_in.Count;
                        Grid_out.ColumnCount = 5;
                        dataGridView3.Columns[0].Visible = true;
                        dataGridView3.Columns[0].Width = 180;
                        dataGridView3.Columns[1].Width = 60;
                        dataGridView3.Columns[2].Width = 60;
                        dataGridView3.Columns[3].Width = 60;
                        dataGridView3.Columns[4].Width = 60;
                        for (int ii = 0; ii < arr_in.Count; ii++)
                        {
                            for (int jj = 0; jj < 5; jj++)
                            {
                                Grid_out.Rows[ii].Cells[jj].Value = String.Format("{0}", arr_in[ii][jj]);
                            }

                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "Сообщение", MessageBoxButtons.OK);
                }
            }
        }


        public void method_of_end_arr(List<List<string>> arr_events_in, List<List<string>> arr_user_in, List<List<string>>  arr_of_deviation_in, ref List<List<string>> arr_out)//метод формирования массива по отработанному времени
        {            
            if (data_is_read == true)
            {

                try
                {
                    List<string> row = new List<string>();
                    arr_out = new List<List<string>>();

                    for (int i = 0; i < arr_user_in.Count; i++)// формирование массива отработанного времени 
                    {
                        row = new List<string>();
                        arr_out.Add(row);
                        arr_out[i].Add("");
                        arr_out[i].Add("");
                        arr_out[i].Add("");
                        arr_out[i].Add("");
                        arr_out[i].Add("");
                        arr_out[i].Add("");
                        arr_out[i][0] = arr_user_in[i][0];
                        arr_out[i][5] = arr_user_in[i][1];//добавление ID пользователя в в 4 столбец,нужен для формирования списка командировка итд
                        for (int ii = 0; ii < arr_events_in.Count; ii++)
                        {
                            if ((arr_user_in[i][2] == arr_events_in[ii][1]) && (Convert.ToInt32(arr_events_in[ii][2]) == 3))
                            {
                                arr_out[i][1] = arr_events_in[ii][0].Remove(0, arr_events_in[ii][0].IndexOf(" ") + 1);
                                                            }
                            if ((arr_user_in[i][2] == arr_events_in[ii][1]) && (Convert.ToInt32(arr_events_in[ii][2]) == 13))
                            {
                                arr_out[i][2] = arr_events_in[ii][0].Remove(0, arr_events_in[ii][0].IndexOf(" ") + 1);

                            }
                        }
                    }
                    for (int i = 0; i < arr_out.Count; i++)//вычисление столбца отработанного времени(8 часов)
                    {

                        if ((arr_out[i][1]!="") && (arr_out[i][2] !=""))// если больше 4 часов на работе то вычитаем час на обед если меньше то нет 
                        {
                            DateTime start = DateTime.Parse(arr_out[i][1]);
                            DateTime end = DateTime.Parse(arr_out[i][2]);
                            if (end >= start)
                            {
                                if(end.Subtract(start)>four_hour)
                                {
                                    arr_out[i][3] = Convert.ToString(end - start - hour);
                                }
                                else
                                {
                                    arr_out[i][3] = Convert.ToString(end - start);
                                }
                            }
                        }
                    }
                    for (int ii = 0; ii < arr_out.Count; ii++)//0 - больничный 1 - отпуск 2 - командировка 3 - удаленная работа
                    {
                        for (int iii = 0; iii < arr_of_deviation_in.Count;iii++)
                        {
                            if (arr_out[ii][5] == arr_of_deviation_in[iii][0])
                            {
                                if (((DateTime.Parse(arr_of_deviation_in[iii][2]) <= DateTime.Parse(date_to_request)) && (DateTime.Parse(date_to_request) <= DateTime.Parse(arr_of_deviation_in[iii][3]))))//проверка на командировку отпуск итд
                                {
                                    arr_out[ii][4] = arr_of_deviation_in[iii][1];
                                    switch (Convert.ToInt16(arr_out[ii][4]))
                                    {
                                        case 0:
                                            arr_out[ii][4] = "больничный";
                                            break;
                                        case 1:
                                            arr_out[ii][4] = "отпуск";
                                            break;
                                        case 2:
                                            arr_out[ii][4] = "командировка";
                                            break;
                                        case 3:
                                            arr_out[ii][4] = "удаленная работа";
                                            break;
                                    }
                                }
                            }

                        }
                    }            

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "Сообщение", MessageBoxButtons.OK);
                }
            }
        }

        public void method_arr_to_excel_eppuls(List<List<string>> arr_in)
        {
            var excelFile = new ExcelPackage();
            var sheet = excelFile.Workbook.Worksheets.Add("Отчет отработанного времени");

            sheet.Cells[1, 1].Value = "Отчет о времени проведенном на рабочем месте за " + date_to_request;
            sheet.Cells[1, 1].Style.Font.Bold = true;
            sheet.Cells[1, 1].Style.Font.Size = 16;
            sheet.Cells["A1:F1"].Merge = true;
            sheet.Column(1).Width = 5;
            sheet.Column(2).Width = 45;
            sheet.Column(3).Width = 14;
            sheet.Column(4).Width = 14;
            sheet.Column(5).Width = 14;
            sheet.Column(6).Width = 14;
            sheet.Cells[2, 1].Value = "№";
            sheet.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells[2, 2].Value = "Фамилия Имя Отчество";
            sheet.Cells[2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[2, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells[2, 3].Value = "Время прихода на работу";
            sheet.Cells[2, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[2, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells[2, 3].Style.WrapText = true;
            sheet.Cells[2, 4].Value = "Время ухода с работы";
            sheet.Cells[2, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[2, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells[2, 4].Style.WrapText = true;
            sheet.Cells[2, 5].Value = "Время нахождения на работе";
            sheet.Cells[2, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[2, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells[2, 5].Style.WrapText = true;
            sheet.Cells[2, 6].Value = "Примечание";
            sheet.Cells[2, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[2, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells[2, 6].Style.WrapText = true;

            for (int i = 0; i < arr_in.Count; i++)
            {
                sheet.Cells[i + 3, 1].Value = i + 1;
                sheet.Cells[i + 3, 2].Value = arr_in[i][0];
                sheet.Cells[i + 3, 3].Value = arr_in[i][1];
                /*if ( (arr_in[i][1] == Convert.ToString("больничный")) || (arr_in[i][1] == Convert.ToString("отпуск")) || (arr_in[i][1] == Convert.ToString("командировка")) || (arr_in[i][1] == Convert.ToString("удаленная работа")))
                {
                    sheet.Cells["C"+(i+3)+":"+"E"+(i+3)].Merge = true;
                }*/
                sheet.Cells[i + 3, 4].Value = arr_in[i][2];
                sheet.Cells[i + 3, 5].Value = arr_in[i][3];
                sheet.Cells[i + 3, 6].Value = arr_in[i][4];
            }

            using (var cells = sheet.Cells[sheet.Dimension.Address])
            {
                cells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //cells.AutoFitColumns();//автоформат ячеек
            }


            var bin = excelFile.GetAsByteArray();
            File.WriteAllBytes(@"Отчет_за_" + date_to_request + ".xlsx", bin);
        }

        public void method_of_deviation(ref List<List<string>> arr_out )//отпуск командировка 0 - больничный 1 - отпуск 2 - командировка 3 - удаленная работа
        {

            try
            {
                if (fb.State == ConnectionState.Open)
                {
                    int i = 0, j = 0;

                    FbTransaction fbt = fb.BeginTransaction();
                    FbCommand SelectSQL = new FbCommand("SELECT deviation.peopleid, deviation.devtype, deviation.devfrom,deviation.devto,deviation.deviationid from deviation", fb); //задаем запрос на получение данных
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
                            arr_out[i].Add("");
                            arr_out[i].Add("");
                            arr_out[i][j] = reader.GetString(0).ToString();//ID пользователя
                            arr_out[i][j + 1] = reader.GetString(1).ToString();//ID состояния
                            int tempInt_in = reader.GetString(2).ToString().IndexOf(" ");
                            string tempT_in2 = reader.GetString(2).ToString().Remove(tempInt_in);//Дата начала 
                            arr_out[i][j + 2] = tempT_in2;
                            int tempInt_out = reader.GetString(3).ToString().IndexOf(" ");
                            string tempT_out2 = reader.GetString(3).ToString().Remove(tempInt_in);//Дата конца
                            arr_out[i][j + 3] = tempT_out2;
                            arr_out[i][j + 4] = reader.GetString(4).ToString();//ID записи 
                            i++;
                        }
                    }
                    finally
                    {
                        //всегда необходимо вызывать метод Close(), когда чтение данных завершено
                        reader.Close();
                        fbt.Dispose();
                        data_is_read = true;
                        //fb.Close(); //закрываем соединение, т.к. оно нам больше не нужно
                    }
                    SelectSQL.Dispose();//в документации написано, что ОЧЕНЬ рекомендуется убивать объекты этого типа, если они больше не нужны
                }
            }
            catch (Exception y)
            {
                MessageBox.Show(y.Message, "Сообщение", MessageBoxButtons.OK);
            }

        }

        //метод формирования массива о отработанном времени в месяц
        public void method_arr_of_events_mounth(Int32 day_end, Int32 current_mounth, Int32 current_year,ref List<List<string>> arr_out)
        {
            try
            {

                bool start_bool = true;
                List<string> row = new List<string>();          
                arr_out = new List<List<string>>();
                row = new List<string>();
                for (int d=1;d<=day_end;d++)
                {
                    Int32 temp_arr_lenth = arr_out.Count;
                    if (fb.State == ConnectionState.Open)
                    {
                        int i = 0, j = 0;
                        FbTransaction fbt = fb.BeginTransaction();
                        FbCommand SelectSQL = new FbCommand("SELECT DISTINCT events.eventsdate, events.cardnum,events.readerid FROM events WHERE events.eventsdate >= '" + d +"." + current_mounth + "." + current_year + " 00:00:00' AND events.eventsdate <= '" + d + "." + current_mounth + "." + current_year + " 23:59:59'", fb); //задаем запрос на выборку
                        SelectSQL.Transaction = fbt;
                        FbDataReader reader = SelectSQL.ExecuteReader();
                        Int32 temp = reader.FieldCount;
                        try
                        {
                            while (reader.Read()) //пока не прочли все данные выполняем...
                            {
                                row = new List<string>();
                                arr_out.Add(row);
                                arr_out[i+ temp_arr_lenth].Add("");
                                arr_out[i+ temp_arr_lenth].Add("");
                                arr_out[i+ temp_arr_lenth].Add("");
                                arr_out[i+ temp_arr_lenth][j] = reader.GetString(0).ToString();//добавиои время
                                arr_out[i+ temp_arr_lenth][j + 1] = reader.GetString(1).ToString();// добавили id ключа
                                arr_out[i+ temp_arr_lenth][j + 2] = reader.GetString(2).ToString();//добавили id эвента
                                for (int ii = temp_arr_lenth; (ii < arr_out.Count - 1) && start_bool == false; ii++)// запуск цикла проверки на существующие запиви
                                {
                                    if ((arr_out[ii][1] == arr_out[i+ temp_arr_lenth][1]) && (arr_out[ii][2] == arr_out[i+ temp_arr_lenth][2]))//проверка на уже существующее ID 
                                    {
                                        if (Convert.ToInt32(arr_out[i+ temp_arr_lenth][2]) == 3)// проверка на сощуствующий эвент 3-вход 13-выход
                                        {
                                            if (DateTime.Parse(arr_out[ii][0]) > DateTime.Parse(arr_out[i+ temp_arr_lenth][0])) // проверка на время
                                            {
                                                arr_out[ii][0] = arr_out[i+ temp_arr_lenth][0];
                                            }
                                        }
                                        if (Convert.ToInt32(arr_out[i+ temp_arr_lenth][2]) == 13)// провоека на сощуствующий эвент 3-вход 13-выход
                                        {
                                            if (DateTime.Parse(arr_out[ii][0]) < DateTime.Parse(arr_out[i+ temp_arr_lenth][0])) // проверка на время
                                            {
                                                arr_out[ii][0] = arr_out[i+ temp_arr_lenth][0];
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
                            fbt.Dispose();
                            //fb.Close(); //закрываем соединение, т.к. оно нам больше не нужно
                        }
                        SelectSQL.Dispose();//в документации написано, что ОЧЕНЬ рекомендуется убивать объекты этого типа, если они больше не нужны
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Сообщение", MessageBoxButtons.OK);
            }

        }

        //метод формирования масива об отработанном времени в течении месяца
               
        public void method_of_end_arr_mounth(List<List<string>> arr_events_in, List<List<string>> arr_user_in, List<List<string>> arr_of_deviation_in, ref List<List<string>> arr_out)//метод формирования массива по отработанному времени
        {        
            Int32 last_day = (DateTime.Parse(arr_events_in[arr_events_in.Count - 1][0])).Day;
            Int32 current_year = (DateTime.Parse(arr_events_in[arr_events_in.Count - 1][0])).Year;
            Int32 current_mount = (DateTime.Parse(arr_events_in[arr_events_in.Count - 1][0])).Month;
            if (data_is_read == true)
            {
                try
                {
                    List<string> row = new List<string>();
                    arr_out = new List<List<string>>();
                    row = new List<string>();
                    arr_out.Add(row);
                    arr_out[0].Add("");
                    arr_out[0].Add("");
                    for (int t=0; t<last_day;t++ )//добавление дат в первую строку
                    {
                        arr_out[0].Add(Convert.ToString(new DateTime(current_year, current_mount, t+1).ToShortDateString()));                        
                    }
                    for (int i = 0; i < arr_user_in.Count; i++)// формирование массива отработанного времени 
                    {
                        
                        row = new List<string>();
                        arr_out.Add(row);
                        for (int t = 0; t < last_day+2; t++)//добавление ячеек по кличетсву дней
                        {
                            arr_out[i+1].Add("");
                        }
                        for (int d = 1; d <= last_day; d++)
                        {
                            var temp_user_i = i;
                            var temp11 = d;
                            DateTime to_work = new DateTime(2000, 1, 1);
                            DateTime from_work = new DateTime(2000, 1, 1);
                            TimeSpan time_in_work;
                            bool find_to_work = false;
                            bool find_from_work = false;
                            arr_out[i+1][0] = arr_user_in[i][1];//ID пользователя в ячейку
                            arr_out[i+1][1] = arr_user_in[i][0];//Имя пользователя в ячейку
                            for (int ii = 0; ii < arr_events_in.Count; ii++)// начало и конец рабочего дня 
                            {
                                if ((arr_user_in[i][2] == arr_events_in[ii][1]) && (Convert.ToInt32(arr_events_in[ii][2]) == 3) && DateTime.Parse(arr_events_in[ii][0]).Day == d)
                                {
                                    to_work = (DateTime.Parse(arr_events_in[ii][0]));
                                    arr_events_in.RemoveAt(ii);
                                    find_to_work = true;
                                }
                                if ((arr_user_in[i][2] == arr_events_in[ii][1]) && (Convert.ToInt32(arr_events_in[ii][2]) == 13) && DateTime.Parse(arr_events_in[ii][0]).Day == d)
                                {
                                    from_work = (DateTime.Parse(arr_events_in[ii][0]));
                                    arr_events_in.RemoveAt(ii);
                                    find_from_work = true;
                                }
                                if (find_to_work && find_from_work)// если время на работе больше 4 часов то 1 час отнимаем если нет то не отнимаем 
                                {
                                    if (from_work.Subtract(to_work) > four_hour)
                                    {
                                        time_in_work = from_work.Subtract(to_work) - hour;
                                        arr_out[i+1][d+1] = Convert.ToString(time_in_work) + " ";
                                    }
                                    else
                                    {
                                        time_in_work = from_work.Subtract(to_work);
                                        arr_out[i+1][d+1] = Convert.ToString(time_in_work)+ " ";
                                    }
                                    find_to_work = false;
                                    find_from_work = false;
                                }
                            }
                        }
                    }
                    for (int ii = 1; ii < arr_out.Count; ii++)//0 - больничный 1 - отпуск 2 - командировка 3 - удаленная работа
                    {
                        for (int iii = 0; iii < arr_of_deviation_in.Count; iii++)
                        {
                            string temp_arr_out = arr_out[ii][0];
                            string temp_arr_dev = arr_of_deviation_in[iii][0];
                            if (arr_out[ii][0] == arr_of_deviation_in[iii][0])
                            {
                                for (int d=0; d<last_day;d++)
                                {
                                    if (((DateTime.Parse(arr_of_deviation_in[iii][2]) <= DateTime.Parse(arr_out[0][d+2])) && (DateTime.Parse(arr_out[0][d+2]) <= DateTime.Parse(arr_of_deviation_in[iii][3]))))
                                    {
                                        string dev_string = "";

                                        switch (Convert.ToInt16(arr_of_deviation_in[iii][1]))
                                        {
                                            case 0:
                                                dev_string = "больничный";
                                                break;
                                            case 1:
                                                dev_string = "отпуск";
                                                break;
                                            case 2:
                                                dev_string = "командировка";
                                                break;
                                            case 3:
                                                dev_string = "удаленная работа";
                                                break;
                                        }
                                        arr_out[ii][d+2] = arr_out[ii][d+2] + "("+dev_string+")";
                                    }
                                }
                            }

                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "Сообщение", MessageBoxButtons.OK);
                }
            }
        }

        public void method_arr_to_grid_mounth(List<List<string>> arr_in, ref DataGridView Grid_out)//метод вывывода массива в датагрид
        {
            if (data_is_read == true)
            {
                try
                {
                    if (arr_in.Count != 0)
                    {
                        int temp = DateTime.Parse(arr_in[0].Last()).Day;
                        Grid_out.RowCount = arr_in.Count;
                        Grid_out.ColumnCount = temp+2;
                        dataGridView3.Columns[0].Visible = false;
                        dataGridView3.Columns[1].Width = 180;
                        for (int ii = 0; ii < arr_in.Count; ii++)
                        {
                            for (int jj = 0; jj < temp+2; jj++)
                            {
                                Grid_out.Rows[ii].Cells[jj].Value = String.Format("{0}", arr_in[ii][jj]);
                            }

                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "Сообщение", MessageBoxButtons.OK);
                }
            }
        }
                     
        public void method_arr_to_excel_mounth_eppuls(List<List<string>> arr_in)//метод записи в файл эексель отчета о отработаном времени 
        {
            var excelFile = new ExcelPackage();
            var sheet = excelFile.Workbook.Worksheets.Add("Отчет отработанного времени в месяц");
            int day_count = DateTime.Parse(arr_in[0].Last()).Day;
            string mounth_string="";
            switch (DateTime.Parse(arr_in[0].Last()).Month)
            {
                case 1:
                    mounth_string = "январь";
                    break;
                case 2:
                    mounth_string = "февраль";
                    break;
                case 3:
                    mounth_string = "март";
                    break;
                case 4:
                    mounth_string = "апрель";
                    break;
                case 5:
                    mounth_string = "май";
                    break;
                case 6:
                    mounth_string = "июнь";
                    break;
                case 7:
                    mounth_string = "июль";
                    break;
                case 8:
                    mounth_string = "август";
                    break;
                case 9:
                    mounth_string = "сентябрь";
                    break;
                case 10:
                    mounth_string = "октябрь";
                    break;
                case 11:
                    mounth_string = "ноябрь";
                    break;
                case 12:
                    mounth_string = "декабрь";
                    break;
            }
            for (int s = 1; s < day_count+1; s++)//утсановливаем размер и перенос строк у ячеек
            {
                sheet.Column(s + 2).Width = 12;
                sheet.Column(s + 2).Style.WrapText = true;
            }

            sheet.Cells[1, 1].Value = "Отчет о времени проведенном на рабочем месте за " + mounth_string + " " + DateTime.Parse(arr_in[0].Last()).Year;//название отчета 
            sheet.Cells[1, 1].Style.Font.Bold = true;
            sheet.Cells[1, 1].Style.Font.Size = 15;
            sheet.Cells["A1:G1"].Merge = true;
            sheet.Column(1).Width = 5;
            sheet.Column(2).Width = 37;

            for (int i = 0; i < arr_in.Count-1; i++)//генерация нумерации
            {
                sheet.Cells[i + 3, 1].Value = i + 1;
            }
            for (int i = 0; i < arr_in.Count; i++)//генерация данных экселя
            {
                for (var j=0; j< day_count+1;j++)
                {
                    sheet.Cells[i + 2, j+2].Value = arr_in[i][j+1];
                }
            }

            sheet.Cells[2, 1].Value = "№";
            sheet.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[2, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells[2, 2].Value = "Фамилия Имя Отчество";
            sheet.Cells[2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells[2, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Row(2).Style.Font.Bold = true;
            sheet.Row(2).Style.Font.Size = 11;

            using (var cells = sheet.Cells[sheet.Dimension.Address])
            {
                cells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }
            var bin = excelFile.GetAsByteArray();
            File.WriteAllBytes(@"Отчет_за_" + mounth_string + "_ " + DateTime.Parse(arr_in[0].Last()).Year + ".xlsx", bin);//запись в фвйл
        }
        
        //------------- Кнопки запуска

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e) //выбор даты запроса
        {
            date_to_request = dateTimePicker1.Value.ToShortDateString();
        }
        
        private void button7_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                connecting_path = method_connection_string(textBox1, textBox2, textBox3);
                method_connect_to_fb(connecting_path, ref label5);// подключаемся к БД
                if (date_to_request == "0")//если 0 то берем сегодняшнюю дату
                {
                    date_to_request = now.ToString("dd.MM.yyyy");
                }
                try
                {
                    method_arr_of_users(ref arr_user);//фомируем массив пользователей
                    method_of_deviation(ref arr_of_deviation);//формируем массив отпусков командировок итд
                    method_arr_of_events(date_to_request, ref arr_events);//формируем массив сообытий
                    method_of_end_arr(arr_events, arr_user, arr_of_deviation, ref arr_of_work);//формируем окончательный массив данных
                    method_arr_to_grid(arr_of_work, ref dataGridView3);//выводим массив в датагрид
                    button3.Enabled = true;
                    button1.Enabled = true;
                }
                catch (Exception r)
                {
                    button3.Enabled = false;
                    button1.Enabled = false;
                    MessageBox.Show(r.Message, "Сообщение", MessageBoxButtons.OK);
                }
            }
            else
            {
                Int32 month = DateTime.Parse(Convert.ToString(dateTimePicker1.Value)).Month;
                Int32 year = DateTime.Parse(Convert.ToString(dateTimePicker1.Value)).Year;
                Int32 day_in_mounth = DateTime.DaysInMonth(year, month);
                try
                {               
                    connecting_path = method_connection_string(textBox1, textBox2, textBox3);
                    method_connect_to_fb(connecting_path, ref label5);// подключаемся к БД
                    method_arr_of_users(ref arr_user);//фомируем массив пользователей
                    method_of_deviation(ref arr_of_deviation);//формируем массив отпусков командировок итд
                    method_arr_of_events_mounth(day_in_mounth, month, year, ref arr_events_per_mounth);//формируем массив отработок за месяц
                    method_of_end_arr_mounth(arr_events_per_mounth, arr_user, arr_of_deviation, ref arr_of_work);//формируем окончательный массив данных
                    method_arr_to_grid_mounth(arr_of_work, ref dataGridView3);//выводим массив в датагрид
                    button3.Enabled = true;
                    button1.Enabled = true;
                }
                catch (Exception r)
                {
                    button3.Enabled = false;
                    button1.Enabled = false;
                    MessageBox.Show(r.Message, "Сообщение", MessageBoxButtons.OK);
                }

            }
        }

        private void button1_Click(object sender, EventArgs e) // кнопка формирование отчетов
        {
            if (radioButton1.Checked)
            {
                try
                {
                    method_arr_to_excel_eppuls(arr_of_work);
                }
                catch (Exception r)
                {
                    MessageBox.Show(r.Message, "Сообщение", MessageBoxButtons.OK);
                }
            }
            else
            {
                try
                {
                    method_arr_to_excel_mounth_eppuls(arr_of_work);
                }
                catch (Exception r)
                {
                    MessageBox.Show(r.Message, "Сообщение", MessageBoxButtons.OK);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form3 fr3 = new Form3();
            fr3.Show();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            Form4 fr4 = new Form4();
            fr4.Show();
        }

        private void radioButton1_MouseClick(object sender, MouseEventArgs e) // выбор отчетов за день
        {
            radioButton1.Checked = true;
            radioButton2.Checked = false;
            dateTimePicker1.Format = DateTimePickerFormat.Short;
        }

        private void radioButton2_MouseClick(object sender, MouseEventArgs e) // выбор отчетов за месяц
        {
            radioButton1.Checked = false;
            radioButton2.Checked = true;
            dateTimePicker1.CustomFormat = "MM.yyyy";
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
        }

        private void button4_Click(object sender, EventArgs e)//кнопка закрытия программы
        {
            Close();
        }        
    }
}
