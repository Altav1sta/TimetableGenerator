using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace TimeTableGenerating
{
    public partial class MainForm : Form
    {
        private TableLayoutPanel[,] tlp;
        private int count;
        private string connectionStr; 

        public MainForm()
        {
            InitializeComponent();

            InitializeTableByGroups();
        }


        private void InitializeTableByGroups()
        {
            // Вычисляем, сколько существует групп
            OleDbConnection con;
            try 
            {
                string tmpStr = string.Format(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}\..\..\sources\TimeTableDB.mdb", Environment.CurrentDirectory);
                con = new OleDbConnection(tmpStr);
                con.Open();
                connectionStr = tmpStr;
            }
            catch (Exception e)
            {
                connectionStr = string.Format(@"Provider=Microsoft.Ace.OLEDB.12.0;Data Source={0}\..\..\sources\TimeTableDB.mdb", Environment.CurrentDirectory);
                con = new OleDbConnection(connectionStr);
                con.Open();
            }
            OleDbCommand oc = new OleDbCommand("SELECT count(*) FROM Groups", con);
            count = Convert.ToInt32(oc.ExecuteScalar());

            // Добавляем столько столбцов, сколько существует групп в базе данных
            this.tlpTable.ColumnCount += count;
            for (int i = 0; i < count; i++)
            {
                this.tlpTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());

            }

            
            oc = new OleDbCommand("SELECT Number FROM Groups", con);
            OleDbDataReader odr = oc.ExecuteReader();
            if (odr.HasRows)
            {
                // Добавляем названия групп в заголовки
                List<int> lst = new List<int>();
                while (odr.Read())
                {
                    lst.Add(odr.GetInt32(0));
                }
                lst.Sort();
                int k = 0;
                foreach (int num in lst)
                {
                    Label lbl = new Label();
                    lbl.Text = num.ToString();
                    lbl.Anchor = System.Windows.Forms.AnchorStyles.None;
                    lbl.AutoSize = true;
                    this.tlpTable.Controls.Add(lbl, 2 + k, 0);
                    k++;
                }


                // и делим ячейки на 5 строк
                tlp = new TableLayoutPanel[5, count];
                for (int i = 0; i < 5; i++)
                {
                    for (int j = 0; j < count; j++)
                    {
                        tlp[i, j] = new TableLayoutPanel();
                        this.tlpTable.Controls.Add(tlp[i, j], 2 + j, 1 + i);

                        tlp[i, j].Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                                    | System.Windows.Forms.AnchorStyles.Left)
                                    | System.Windows.Forms.AnchorStyles.Right)));
                        tlp[i, j].AutoSize = true;
                        tlp[i, j].CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
                        tlp[i, j].ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
                        tlp[i, j].Margin = new System.Windows.Forms.Padding(0);
                        tlp[i, j].RowCount = 5;
                        tlp[i, j].RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
                        tlp[i, j].RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
                        tlp[i, j].RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
                        tlp[i, j].RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
                        tlp[i, j].RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20F));
                    }
                }
            }


            con.Close();

        }


        // --- Event Handlers

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            double[] w = new double[6];

            if (!Double.TryParse(textBox1.Text, out w[0]) || !Double.TryParse(textBox2.Text, out w[1]) || !Double.TryParse(textBox3.Text, out w[2]) || !Double.TryParse(textBox4.Text, out w[3])
                || !Double.TryParse(textBox5.Text, out w[4]) || !Double.TryParse(textBox6.Text, out w[5]))
            {
                MessageBox.Show("Значение должно быть десятичным числом!");
                return;
            }

            Generator g = new Generator(count, w, connectionStr);
            g.generateTimeTable();
            

            for (int day = 0; day < 5; day++)
            {
                for (int gr = 0; gr < count; gr++)
                {
                    tlp[day, gr].Controls.Clear();

                    for (int pos = 0; pos < 5; pos++)
                    {
                        Label lbl = new Label();
                        if (g.timetable[gr, day, pos].id == 0) lbl.Text = "";
                        else lbl.Text = g.timetable[gr, day, pos].toString(connectionStr);
                        lbl.Anchor = System.Windows.Forms.AnchorStyles.None;
                        lbl.AutoSize = true;

                        tlp[day, gr].Controls.Add(lbl, 0, pos);
                    }
                }
            }
            button1.Enabled = true;
        }
    }
}
