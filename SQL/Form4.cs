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
        Int32 Current_Row = 0;

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


        public void method_DataGridDeviation(bool need_to_lock ,ref CheckBox checkbox_in, ref DataGridView datagrid_of_dev)//вывод данных в датагрид
        {
            datagrid_of_dev.ReadOnly = true;
            datagrid_of_dev.RowsDefaultCellStyle.BackColor = Color.LightGray;
            checkbox_in.Checked = false;
            datagrid_of_dev.Columns.Clear();
            DataGridViewComboBoxColumn boxcolum = new DataGridViewComboBoxColumn();
            boxcolum.HeaderText = "Причина отсуствия";
            boxcolum.DropDownWidth = 90;
            boxcolum.Width = 90;
            boxcolum.MaxDropDownItems = 4;
            datagrid_of_dev.Columns.Insert(0, boxcolum);
            boxcolum.Items.AddRange("больничный", "отпуск", "командировка", "удаленная работа");
            datagrid_of_dev.Rows.Clear();
            datagrid_of_dev.ColumnCount = 4;
            datagrid_of_dev.Columns[1].Width = 90;
            datagrid_of_dev.Columns[1].HeaderText = "Начальная дата";
            datagrid_of_dev.Columns[2].Width = 90;
            datagrid_of_dev.Columns[2].HeaderText = "Конечная дата";
            datagrid_of_dev.Columns[3].Width = 10;
            datagrid_of_dev.Columns[3].Visible = false;
            User_ID = Program.f1.arr_user.Where(o => o.IndexOf(listBox1.SelectedItem.ToString()) != -1).FirstOrDefault()[1];// полчение ID выбранного пользователя(выбор+поиск по массиву пользователей)
            var temp = Program.f1.arr_of_deviation.Where(o => o[0] == User_ID).ToList();

            for (int i = 0; i < temp.Count; i++)
            {

                datagrid_of_dev.Rows.Add();
                switch (Convert.ToInt16(temp[i][1]))
                {
                    case 0:
                        datagrid_of_dev.Rows[i].Cells[0].Value = "больничный";
                        break;
                    case 1:
                        datagrid_of_dev.Rows[i].Cells[0].Value = "отпуск";
                        break;
                    case 2:
                        datagrid_of_dev.Rows[i].Cells[0].Value = "командировка";
                        break;
                    case 3:
                        datagrid_of_dev.Rows[i].Cells[0].Value = "удаленная работа";
                        break;
                }
                datagrid_of_dev.Rows[i].Cells[1].Value = temp[i][2];
                datagrid_of_dev.Rows[i].Cells[2].Value = temp[i][3];
                datagrid_of_dev.Rows[i].Cells[3].Value = temp[i][4];
            }

            if (need_to_lock == true)
            {
                datagrid_of_dev.ReadOnly = true;
                datagrid_of_dev.RowsDefaultCellStyle.BackColor = Color.LightGray;
            }
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e) // заполнение датагрида 
        {
            method_DataGridDeviation(true,ref checkBox1,ref dataGridView1);
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (checkBox1.Checked)
            {
                
                if (Current_Row == 0 || Current_Row == e.RowIndex)
                {
                    
                    if (fb.State != ConnectionState.Open)
                    { method_connect_to_fb(Program.f1.connecting_path); }

                    if (dataGridView1.RowCount - 1 == e.RowIndex)//добавление новой строки в БД
                    {
                        dataGridView1.AllowUserToAddRows = false;
                        checkBox1.Enabled = false;
                        need_to_end_new_line = true;
                        label3.Visible = true;
                        int reason_absence = -1;
                        label3.ForeColor = Color.Red;
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

                        if (Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[0].Value) == "больничный")
                        {
                            reason_absence = 0;
                        }

                        if (Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[0].Value) == "отпуск")
                        {
                            reason_absence = 1;
                        }
                        if (Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[0].Value) == "командировка")
                        {
                            reason_absence = 2;
                        }
                        if (Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[0].Value) == "удаленная работа")
                        {
                            reason_absence = 3;
                        }

                        if ((Regex.IsMatch(Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[1].Value), pattern) && Regex.IsMatch(Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[2].Value), pattern)) && reason_absence != (-1))
                        {
                            try
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[2].Style.BackColor = Color.White;
                                FbCommand InsertSQL = new FbCommand("insert into deviation(deviation.peopleid,deviation.devfrom,deviation.devto,deviation.devtype) values('" + User_ID + "','" + dataGridView1.Rows[e.RowIndex].Cells[1].Value + "','" + dataGridView1.Rows[e.RowIndex].Cells[2].Value + "','" + reason_absence + "')", fb); //задаем запрос вовод данных
                                if (fb.State == ConnectionState.Open)
                                {
                                    FbTransaction fbt = fb.BeginTransaction(); //необходимо проинициализить транзакцию для объекта InsertSQL
                                    InsertSQL.Transaction = fbt;
                                    int result = InsertSQL.ExecuteNonQuery();
                                    //MessageBox.Show("Добавление причины отсутвия на рабочем месте выполнено");
                                    fbt.Commit();
                                    fbt.Dispose();
                                    InsertSQL.Dispose();
                                    need_to_end_new_line = false;
                                    label3.Visible = false;
                                    for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                                    {
                                        dataGridView1.Rows[i].Cells[0].ReadOnly = false;
                                        dataGridView1.Rows[i].Cells[1].ReadOnly = false;
                                        dataGridView1.Rows[i].Cells[2].ReadOnly = false;
                                    }
                                    checkBox1.Enabled = true;
                                    Program.f1.method_of_deviation(ref Program.f1.arr_of_deviation);
                                    this.BeginInvoke(new MethodInvoker(() =>
                                    {
                                        method_DataGridDeviation(false, ref checkBox1, ref dataGridView1);//запуск асинхронного метода
                                    }));
                                    //method_DataGridDeviation(false, ref checkBox1, ref dataGridView1); // необходимо решить проблему с вылетом когда при заполнении выбирается другая ячейка
                                    dataGridView1.AllowUserToAddRows = true;
                                    Current_Row = 0;
                                }
                            }
                            catch (Exception r)
                            {
                                MessageBox.Show(r.Message, "Сообщение", MessageBoxButtons.OK);
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
                            try
                            {
                                FbCommand InsertSQL = new FbCommand("update deviation set deviation.devfrom='" + dataGridView1.Rows[e.RowIndex].Cells[1].Value + "', deviation.devto='" + dataGridView1.Rows[e.RowIndex].Cells[2].Value + "', deviation.devtype='" + reason_absence + "'where deviation.deviationid='" + dataGridView1.Rows[e.RowIndex].Cells[3].Value + "'", fb); //задаем запрос на получение данных
                                if (fb.State == ConnectionState.Open)
                                {
                                    FbTransaction fbt = fb.BeginTransaction(); //необходимо проинициализить транзакцию для объекта InsertSQL
                                    InsertSQL.Transaction = fbt;
                                    int result = InsertSQL.ExecuteNonQuery();
                                    //MessageBox.Show("Изменение причины отсутвия на рабочем месте выполнено");
                                    fbt.Commit();
                                    fbt.Dispose();
                                    InsertSQL.Dispose();
                                    label3.Visible = false;
                                    for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                                    {
                                        dataGridView1.Rows[i].Cells[0].ReadOnly = false;
                                        dataGridView1.Rows[i].Cells[1].ReadOnly = false;
                                        dataGridView1.Rows[i].Cells[2].ReadOnly = false;
                                    }
                                    checkBox1.Enabled = true;
                                    Program.f1.method_of_deviation(ref Program.f1.arr_of_deviation);
                                    this.BeginInvoke(new MethodInvoker(() =>
                                    {
                                        method_DataGridDeviation(false, ref checkBox1, ref dataGridView1);// запуск асинхронного метода
                                    }
                                    ));
                                    //method_DataGridDeviation(false, ref checkBox1, ref dataGridView1);// необходимо решить проблему с вылетом когда при заполнении выбирается другая ячейка
                                    dataGridView1.AllowUserToAddRows = true;
                                    Current_Row = 0;
                                }
                            }
                            catch(Exception r)
                            {
                                MessageBox.Show(r.Message, "Сообщение", MessageBoxButtons.OK);
                            }
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
