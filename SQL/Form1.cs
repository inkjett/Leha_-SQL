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
using OfficeOpenXml;
using OfficeOpenXml.Style;

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
        public bool data_is_read = false;
        public string path_to_DB = "C:\\123.fdb";
        public string User = "SYSDBA";
        public string Pass = "masterkey";
        DateTime now = DateTime.Now;
        Form3 fr3 = new Form3();
        TimeSpan infinite = TimeSpan.FromMilliseconds(-1);
        TimeSpan hour = new TimeSpan(1, 0, 0);



        //------
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }


        //методы
        public void method_connect_to_fb(TextBox Text_Path, TextBox Text_user, TextBox Text_pass, ref Label Label_out)// метод подключения к БД
        {
            try
            {
                string path_path = "character set = WIN1251; initial catalog = "+textBox4.Text+":" + @""+textBox1.Text+"; user id = " + Text_user.Text + "; password = " + Text_pass.Text + "; ";
                fb = new FbConnection(path_path);
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
                        Grid_out.ColumnCount = 4;
                        dataGridView3.Columns[0].Width = 180;
                        dataGridView3.Columns[1].Width = 60;
                        dataGridView3.Columns[2].Width = 60;
                        dataGridView3.Columns[3].Width = 60;
                        for (int ii = 0; ii < arr_in.Count; ii++)
                        {
                            for (int jj = 0; jj < 4; jj++)
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


        public void method_of_end_arr(List<List<string>> arr_events_in, List<List<string>> arr_user_in, ref List<List<string>> arr_out)//метод формирования массива по отработанному времени
        {

            if (data_is_read == true)
            {

                try
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
                        arr_out[i].Add("");
                        arr_out[i][0] = arr_user_in[i][0];

                        for (int ii = 0; ii < arr_events_in.Count; ii++)// формирование массива 
                        {
                            if ((arr_user_in[i][2] == arr_events_in[ii][1]) && (Convert.ToInt32(arr_events_in[ii][2]) == 3))
                            {
                                string temp = arr_events_in[ii][0];
                                int value_of_index = temp.IndexOf(" ");
                                string temp2 = temp.Remove(0, value_of_index+1);
                                arr_out[i][1] = temp2;
                            }
                            if ((arr_user_in[i][2] == arr_events_in[ii][1]) && (Convert.ToInt32(arr_events_in[ii][2]) == 13))
                            {
                                string temp = arr_events_in[ii][0];
                                int value_of_index = temp.IndexOf(" ");
                                string temp2 = temp.Remove(0, value_of_index+1);
                                arr_out[i][2] = temp2;
                            }
                        }

                    }

                    for (int i = 0; i < arr_out.Count; i++)
                    {

                        if ((arr_out[i][1]!="") && (arr_out[i][2] !=""))
                        {
                            DateTime start = DateTime.Parse(arr_out[i][1]);
                            DateTime end = DateTime.Parse(arr_out[i][2]);
                            if (end >= start)
                            {
                                arr_out[i][3] = Convert.ToString(end - start - hour);
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
            sheet.Cells["A1:E1"].Merge = true;
            sheet.Column(1).Width = 5;
            sheet.Column(2).Width = 45;
            sheet.Column(3).Width = 12;
            sheet.Column(4).Width = 12;
            sheet.Column(5).Width = 12;
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


            for (int i = 0; i < arr_in.Count; i++)
            {
                sheet.Cells[i + 3, 1].Value = i + 1;
                sheet.Cells[i + 3, 2].Value = arr_in[i][0];
                sheet.Cells[i + 3, 3].Value = arr_in[i][1];
                sheet.Cells[i + 3, 4].Value = arr_in[i][2];
                sheet.Cells[i + 3, 5].Value = arr_in[i][3];
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
 
               
        //-------------


        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            date_to_request = dateTimePicker1.Value.ToShortDateString();
        }


        private void button7_Click(object sender, EventArgs e)
        {

            method_connect_to_fb(textBox1, textBox2, textBox3, ref label5);
            if (date_to_request == "0")
            {
                date_to_request = now.ToString("dd.MM.yyyy");
            }
            try
            {
                method_arr_of_users(ref arr_user);
                method_arr_of_events(date_to_request, ref arr_events);
                method_of_end_arr(arr_events, arr_user, ref arr_of_work);
                method_arr_to_grid(arr_of_work, ref dataGridView3);
            }
            catch (Exception r)
            {
                MessageBox.Show(r.Message, "Сообщение", MessageBoxButtons.OK);
            }

        }
        
        private void button4_Click(object sender, EventArgs e)
        {
            Close(); 
        }

        private void button1_Click(object sender, EventArgs e)
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

        private void button2_Click(object sender, EventArgs e)
        {
            fr3.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "DB (*.fdb)|*.fdb";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

    }
}
