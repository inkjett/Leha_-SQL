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
            textBox4.Text = Form1.IP;
            textBox1.Text = Form1.pathToDB;
            textBox2.Text = Form1.User;
            textBox3.Text = Form1.Password;
        }
        Form1 Form_item = new Form1();
        private void button1_Click(object sender, EventArgs e)
        {
            OPTsettings.Props.writteXML();
            //OPTsettings.Props.readerXML();
        }

    }
}
