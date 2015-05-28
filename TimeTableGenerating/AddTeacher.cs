using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TimeTableGenerating
{
    public partial class AddTeacher : Form
    {
        public string query;

        public AddTeacher()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string tmpStr = textBox3.Text.Trim();
            if (!tmpStr.Equals(""))
            {
                query = "INSERT INTO Teachers (TName, Monday, Tuesday, Wednesday, Thursday, Friday) values('" + tmpStr + "', " + checkBox5.Checked + ", " + checkBox4.Checked + ", " + checkBox3.Checked
                    + ", " + checkBox2.Checked + ", " + checkBox1.Checked + ")";

                textBox3.Text = "";
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
            }
        }


    }
}
