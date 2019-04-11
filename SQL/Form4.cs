using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SQL
{
    public partial class Form4 : Form
    {


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

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            var temp = e.RowIndex;
            var temp2 = e.ColumnIndex;
        }
    }
}
