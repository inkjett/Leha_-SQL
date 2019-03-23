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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        Form1 Form_item = new Form1();
        private void button1_Click(object sender, EventArgs e)
        {
            Form_item.path_to_DB = textBox1.Text;
            Form_item.User = textBox2.Text;
            Form_item.path_to_DB = textBox3.Text;            
        }
    }
}
