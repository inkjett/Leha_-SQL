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

        public Form4()
        {
            InitializeComponent();
        }
        List<List<string>> arr_of_deviation_to_DB;
        private void button1_Click(object sender, EventArgs e)
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
                //string path_path = "character set = WIN1251; initial catalog = "+textBox4.Text+":" + @""+textBox1.Text+"; user id = " + Text_user.Text + "; password = " + Text_pass.Text + "; ";
                fb = new FbConnection(path_in);
                fb.Open();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Сообщение", MessageBoxButtons.OK);
            }
        }





        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int GridCount = 0;
            dataGridView1.Rows.Clear();
            dataGridView1.ColumnCount = 4;
            dataGridView1.Columns[0].Width = 90;
            dataGridView1.Columns[1].Width = 90;
            dataGridView1.Columns[2].Width = 90;
            dataGridView1.Columns[3].Width = 1;
            dataGridView1.Columns[3].Visible = false;
            string User_ID = Program.f1.arr_user.Where(o => o.IndexOf(listBox1.SelectedItem.ToString()) != -1).FirstOrDefault()[1];// полчение ID выбранного пользователя(выбор+поиск по массиву пользователей)
            var temp = Program.f1.arr_of_deviation.Where(o => o[0] == User_ID).ToList();
            var temp2 = temp;

            for (int i=0; i<temp.Count;i++)
            {
                dataGridView1.Rows.Add();
                switch (Convert.ToInt16(temp[i][1]))
                {
                    case 0:
                        dataGridView1.Rows[i].Cells[0].Value = "больничный";
                        break;
                    case 1:
                        dataGridView1.Rows[i].Cells[0].Value = "отпуск";
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

        private void button2_Click(object sender, EventArgs e)
        {

            //var temp = e.RowIndex;
            //var temp2 = e.ColumnIndex;
            method_connect_to_fb(Program.f1.connecting_path);
            //string QerySQL = "update deviation set deviation.devfrom='" + "02.04.2020  0:00:00" + "', deviation.devto='" + "14.04.2020  0:00:00" + "', deviation.devtype=0 where deviation.deviationid='12'";
            FbCommand InsertSQL = new FbCommand("update deviation set deviation.devfrom='" + "02.04.2020  0:00:00" + "', deviation.devto='" + "14.04.2020  0:00:00" + "', deviation.devtype=0 where deviation.deviationid='12'", fb); //задаем запрос на получение данных
            
            if (fb.State == ConnectionState.Open)
            {
                FbTransaction fbt = fb.BeginTransaction(); //необходимо проинициализить транзакцию для объекта InsertSQL
                InsertSQL.Transaction = fbt;
                int result = InsertSQL.ExecuteNonQuery();
                MessageBox.Show("ВЫполнено" + result);
                fbt.Commit();
                fbt.Dispose();
                InsertSQL.Dispose();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            string datagrid

            int reason_absence = -1;
            bool find_reason = false;

            string pattern = @"^[0-3]{1}[0-9]{1}.[0-1]{1}[0-9]{1}.[2]{1}[0-1]{1}[0-9]{2}$";
            method_connect_to_fb(Program.f1.connecting_path);
            if(Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[0].Value) == "больничный")
            {
                reason_absence = 0;
                find_reason = true;

            }
            if (Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[0].Value) == "отпуск")
            {
                reason_absence = 1;
                find_reason = true;
            }
            if (Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[0].Value) == "командировка")
            {
                reason_absence = 2;
                find_reason = true;
            }
            if (Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[0].Value) == "удаленная работа")
            {
                reason_absence = 3;
                find_reason = true;
            }
            

            if ((Regex.IsMatch(Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value), pattern)))
            {
                FbCommand InsertSQL = new FbCommand("update deviation set deviation.devfrom='" + dataGridView1.Rows[e.RowIndex].Cells[1].Value + "', deviation.devto='" + dataGridView1.Rows[e.RowIndex].Cells[2].Value + "', deviation.devtype='" + reason_absence +  "'where deviation.deviationid='" + dataGridView1.Rows[e.RowIndex].Cells[3].Value + "'", fb); //задаем запрос на получение данных
                if (fb.State == ConnectionState.Open)
                {
                    FbTransaction fbt = fb.BeginTransaction(); //необходимо проинициализить транзакцию для объекта InsertSQL
                    InsertSQL.Transaction = fbt;
                    int result = InsertSQL.ExecuteNonQuery();
                    MessageBox.Show("ВЫполнено" + result);
                    fbt.Commit();
                    fbt.Dispose();
                    InsertSQL.Dispose();
                }
            }
            else
            {
                MessageBox.Show("Проверьте введенные данные");
            }




        }
        
    }
}
