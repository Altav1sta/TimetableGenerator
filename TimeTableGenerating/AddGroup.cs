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
    public partial class AddGroup : Form
    {
        public string query;
        public AddGroup()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int tmp;
            if (!textBox3.Text.Trim().Equals("") && !textBox4.Text.Trim().Equals("") && Int32.TryParse(textBox3.Text.Trim(), out tmp) && (tmp > 0)
                && Int32.TryParse(textBox4.Text.Trim(), out tmp) && (tmp > 0))
            {
                query = "INSERT INTO Groups values('" + textBox3.Text.Trim() + "', '" + textBox4.Text.Trim() + "')";

                textBox3.Text = "";
                textBox4.Text = "";
            }
        }
    }
}
