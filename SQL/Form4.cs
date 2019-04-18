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
using System.Text.RegularExpressions;


namespace SQL
{
    public partial class Form4 : Form
    {
        FbConnection fb;
        string User_ID = "";//текущий пользователь
        string pattern = @"^[0-3]{1}[0-9]{1}.[0-1]{1}[0-9]{1}.[2]{1}[0-1]{1}[0-9]{2}$";//строка дял проверки ввода даты на формат XX.XX.XXXX                   
        bool need_to_end_new_line = false;
        public Form4()

        {
            InitializeComponent();
            method_connect_to_fb(Program.f1.connecting_path);
        }

        private void Form4_Load(object sender, EventArgs e)//загрузка формы - загрука из другой формы массива о состояниях
        {
            for (int i = 0; i < Program.f1.arr_user.Count; i++)
            {
                int index = Convert.ToInt32(Program.f1.arr_user[i][1]);
                listBox1.Items.Insert(i, Program.f1.arr_user[i][0]);
            }
        }

        public void method_connect_to_fb(string path_in)// метод подключения к БД
        {
            try
            {                
                fb = new FbConnection(path_in);
                fb.Open();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Сообщение", MessageBoxButtons.OK);
            }
        }                     

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e) // заполнение датагрида 
        {
            checkBox1.Checked = false;
            dataGridView1.Columns.Clear();
            DataGridViewComboBoxColumn boxcolum = new DataGridViewComboBoxColumn();
            boxcolum.HeaderText = "Причина отсуствия";
            boxcolum.DropDownWidth = 90;
            boxcolum.Width = 90;
            boxcolum.MaxDropDownItems = 4;
            this.dataGridView1.Columns.Insert(0, boxcolum);
            boxcolum.Items.AddRange("больничный", "отпуск", "командировка", "удаленная работа");            
            dataGridView1.Rows.Clear();
            dataGridView1.ColumnCount = 4;
            dataGridView1.Columns[1].Width = 90;
            dataGridView1.Columns[1].HeaderText = "Начальная дата";
            dataGridView1.Columns[2].Width = 90;
            dataGridView1.Columns[2].HeaderText = "Конечная дата";
            dataGridView1.Columns[3].Width = 1;
            dataGridView1.Columns[3].Visible = false;      
            dataGridView1.ReadOnly = true;
            dataGridView1.RowsDefaultCellStyle.BackColor = Color.LightGray;
            User_ID = Program.f1.arr_user.Where(o => o.IndexOf(listBox1.SelectedItem.ToString()) != -1).FirstOrDefault()[1];// полчение ID выбранного пользователя(выбор+поиск по массиву пользователей)
            var temp = Program.f1.arr_of_deviation.Where(o => o[0] == User_ID).ToList();

            for (int i=0; i<temp.Count;i++)
            {
                
                dataGridView1.Rows.Add();
                switch (Convert.ToInt16 (temp[i][1]))
                {
                    case 0:
                        dataGridView1.Rows[i].Cells[0].Value = "больничный";
                        break;
                    case 1:
                        dataGridView1.Rows[i].Cells[0].Value =  "отпуск";
                        break;
                    case 2:
                        dataGridView1.Rows[i].Cells[0].Value = "командировка";
                        break;
                    case 3:
                        dataGridView1.Rows[i].Cells[0].Value = "удаленная работа";
                        break;
                }                
                dataGridView1.Rows[i].Cells[1].Value = temp[i][2];
                dataGridView1.Rows[i].Cells[2].Value = temp[i][3];
                dataGridView1.Rows[i].Cells[3].Value = temp[i][4];                
            }
            
                                                                              
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            Int32 temp_rowCount = 0;
            if (checkBox1.Checked)
            {                                                                      
                if (dataGridView1.AllowUserToAddRows == false)
                {
                    temp_rowCount = 1;
                }
                else
                {
                    temp_rowCount = 2;
                }
                if (fb.State != ConnectionState.Open)
                { method_connect_to_fb(Program.f1.connecting_path); }
                
                if (dataGridView1.RowCount - temp_rowCount == e.RowIndex)//добавление новой строки в БД
                {
                    checkBox1.Enabled = false;
                    need_to_end_new_line = true;
                    label3.Visible = true;
                    label3.ForeColor = Color.Red;
                    dataGridView1.AllowUserToAddRows = false;
                    label3.Text = "Необходимо завершить ввод/изменение причины отсутвия на рабочем месте";


                    for (int i = 0; i < dataGridView1.RowCount - 1; i++)//ограничение ввода в процессе заполненя новой стороки 
                    {
                        dataGridView1.Rows[i].Cells[0].ReadOnly = true;
                        dataGridView1.Rows[i].Cells[1].ReadOnly = true;
                        dataGridView1.Rows[i].Cells[2].ReadOnly = true;
                    }
                    
                    if (!Regex.IsMatch(Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[1].Value), pattern))//подсветка ячеек при неверном вводе
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[1].Style.BackColor = Color.Red;
                    }
                    else
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[1].Style.BackColor = Color.White;
                    }
                    if (!Regex.IsMatch(Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[2].Value), pattern))
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[2].Style.BackColor = Color.Red;
                    }
                    else
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[2].Style.BackColor = Color.White;
                    }

                    if (Regex.IsMatch(Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[1].Value), pattern) && Regex.IsMatch(Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[2].Value), pattern))
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[2].Style.BackColor = Color.White;
                        FbCommand InsertSQL = new FbCommand("insert into deviation(deviation.peopleid,deviation.devfrom,deviation.devto,deviation.devtype) values('" + User_ID + "','" + dataGridView1.Rows[e.RowIndex].Cells[1].Value + "','" + dataGridView1.Rows[e.RowIndex].Cells[2].Value + "','2')", fb); //задаем запрос на получение данных
                        if (fb.State == ConnectionState.Open)
                        {
                            FbTransaction fbt = fb.BeginTransaction(); //необходимо проинициализить транзакцию для объекта InsertSQL
                            InsertSQL.Transaction = fbt;
                            int result = InsertSQL.ExecuteNonQuery();
                            MessageBox.Show("Добавление причины отсутвия на рабочем месте выполнено");
                            fbt.Commit();
                            fbt.Dispose();
                            InsertSQL.Dispose();
                            need_to_end_new_line = false;
                            label3.Visible = false;
                            dataGridView1.AllowUserToAddRows = true;
                            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                            {
                                dataGridView1.Rows[i].Cells[0].ReadOnly = false;
                                dataGridView1.Rows[i].Cells[1].ReadOnly = false;
                                dataGridView1.Rows[i].Cells[2].ReadOnly = false;
                            }
                            checkBox1.Enabled = true;
                            Program.f1.method_of_deviation(ref Program.f1.arr_of_deviation);
                        }
                    }
                }
                else if (!need_to_end_new_line)//изменение текущей строки в БД
                {
                    int reason_absence = -1;
                    bool can_run_query = false;
                    checkBox1.Enabled = false;
                    dataGridView1.AllowUserToAddRows = false;
                    label3.Visible = true;
                    label3.ForeColor = Color.Red;
                    label3.Text = "Необходимо завершить ввод/изменение причины отсутвия на рабочем месте";
                    for (int i = 0; i < dataGridView1.RowCount - 1; i++)//ограничение ввода в процессе заполненя новой стороки 
                    {
                        dataGridView1.Rows[i].Cells[0].ReadOnly = true;
                        dataGridView1.Rows[i].Cells[1].ReadOnly = true;
                        dataGridView1.Rows[i].Cells[2].ReadOnly = true;
                    }

                    if (Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[0].Value) == "больничный")
                    {
                        reason_absence = 0;
                        can_run_query = true;
                    }

                    if (Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[0].Value) == "отпуск")
                    {
                        reason_absence = 1;
                        can_run_query = true;
                    }
                    if (Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[0].Value) == "командировка")
                    {
                        reason_absence = 2;
                        can_run_query = true;
                    }
                    if (Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[0].Value) == "удаленная работа")
                    {
                        reason_absence = 3;
                        can_run_query = true;
                    }

                    if ((Regex.IsMatch(Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[1].Value), pattern)) && can_run_query)
                    {
                        can_run_query = true;
                        dataGridView1.Rows[e.RowIndex].Cells[1].Style.BackColor = Color.White;
                    }
                    else
                    {
                        can_run_query = false;
                        dataGridView1.Rows[e.RowIndex].Cells[1].Style.BackColor = Color.Red;
                    }

                    if ((Regex.IsMatch(Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[2].Value), pattern)) && can_run_query)
                    {
                        can_run_query = true;
                        dataGridView1.Rows[e.RowIndex].Cells[2].Style.BackColor = Color.White;
                    }
                    else
                    {
                        can_run_query = false;
                        dataGridView1.Rows[e.RowIndex].Cells[2].Style.BackColor = Color.Red;
                    }
                    
                    if (can_run_query)
                    {
                        FbCommand InsertSQL = new FbCommand("update deviation set deviation.devfrom='" + dataGridView1.Rows[e.RowIndex].Cells[1].Value + "', deviation.devto='" + dataGridView1.Rows[e.RowIndex].Cells[2].Value + "', deviation.devtype='" + reason_absence + "'where deviation.deviationid='" + dataGridView1.Rows[e.RowIndex].Cells[3].Value + "'", fb); //задаем запрос на получение данных
                        if (fb.State == ConnectionState.Open)
                        {
                            FbTransaction fbt = fb.BeginTransaction(); //необходимо проинициализить транзакцию для объекта InsertSQL
                            InsertSQL.Transaction = fbt;
                            int result = InsertSQL.ExecuteNonQuery();
                            MessageBox.Show("Изменение причины отсутвия на рабочем месте выполнено");
                            fbt.Commit();
                            fbt.Dispose();
                            InsertSQL.Dispose();
                            label3.Visible = false;
                            dataGridView1.AllowUserToAddRows = true;
                            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                            {
                                dataGridView1.Rows[i].Cells[0].ReadOnly = false;
                                dataGridView1.Rows[i].Cells[1].ReadOnly = false;
                                dataGridView1.Rows[i].Cells[2].ReadOnly = false;
                            }
                            checkBox1.Enabled = true;
                            Program.f1.method_of_deviation(ref Program.f1.arr_of_deviation);
                        }
                    }
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked==true)
            {
                dataGridView1.ReadOnly = false;
                dataGridView1.RowsDefaultCellStyle.BackColor = Color.White;
            }
            else
            {
                dataGridView1.ReadOnly = true;
                dataGridView1.RowsDefaultCellStyle.BackColor = Color.LightGray; 
            }
        }
    }
}
