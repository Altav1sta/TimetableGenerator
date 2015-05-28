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
using Excel = Microsoft.Office.Interop.Excel;

namespace TimeTableGenerating
{
    public partial class MainForm : Form
    {
        private TableLayoutPanel[,] tlp;
        private int count;
        private string connectionStr;
        private List<int> groupList;
        Generator gen;

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
                groupList = new List<int>();
                while (odr.Read())
                {
                    groupList.Add(odr.GetInt32(0));
                }
                groupList.Sort();
                int k = 0;
                foreach (int num in groupList)
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

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void saveInExcel(Generator g)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            Excel.Range chartRange;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            // Номера групп
            xlWorkSheet.Cells[1, 1] = "";
            xlWorkSheet.Cells[1, 2] = "";
            int tmp = 3;
            foreach (int gr in groupList)
            {
                xlWorkSheet.Cells[1, tmp] = gr.ToString();
                tmp++;
            }

            // Номера пар
            for (int day = 0; day < 5; day++)
            {
                xlWorkSheet.Cells[2 + day * 5, 2] = "1 пара";
                xlWorkSheet.Cells[3 + day * 5, 2] = "2 пара";
                xlWorkSheet.Cells[4 + day * 5, 2] = "3 пара";
                xlWorkSheet.Cells[5 + day * 5, 2] = "4 пара";
                xlWorkSheet.Cells[6 + day * 5, 2] = "5 пара";                
            }

            // Занятия
            for (int day = 0; day < 5; day++)
            {
                for (int gr = 0; gr < count; gr++)
                {
                    
                    for (int pos = 0; pos < 5; pos++)
                    {
                        if (g.timetable[gr, day, pos].id == 0) xlWorkSheet.Cells[2 + day * 5 + pos, 3 + gr] = "";
                        else xlWorkSheet.Cells[2 + day * 5 + pos, 3 + gr] = g.timetable[gr, day, pos].toString(connectionStr);
                    }
                }
            }

            // Дни недели
            xlWorkSheet.get_Range("a2", "a6").Merge(false);
            chartRange = xlWorkSheet.get_Range("a2", "a6");
            chartRange.FormulaR1C1 = "Пн";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            
            xlWorkSheet.get_Range("a7", "a11").Merge(false);
            chartRange = xlWorkSheet.get_Range("a7", "a11");
            chartRange.FormulaR1C1 = "Вт";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            xlWorkSheet.get_Range("a12", "a16").Merge(false);
            chartRange = xlWorkSheet.get_Range("a12", "a16");
            chartRange.FormulaR1C1 = "Ср";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            xlWorkSheet.get_Range("a17", "a21").Merge(false);
            chartRange = xlWorkSheet.get_Range("a17", "a21");
            chartRange.FormulaR1C1 = "Чт";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            xlWorkSheet.get_Range("a22", "a26").Merge(false);
            chartRange = xlWorkSheet.get_Range("a22", "a26");
            chartRange.FormulaR1C1 = "Пт";
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 3;
            chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            


            // Рамки
            for (int i = 0; i < 5; i++)
            {
                chartRange = xlWorkSheet.get_Range(string.Format("b{0}", (2+5*i).ToString()), string.Format("b{0}", (6+5*i).ToString()));
                chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            }
            for (int i = 0; i < count; i++)
            {
                // Вокруг групп
                chartRange = xlWorkSheet.get_Range(string.Format("{0}1", (char)(99 + i)), string.Format("{0}1", (char)(99 + i)));
                chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 3;

                // Вокруг расписания группы
                for (int j = 0; j < 5; j++)
                {
                    chartRange = xlWorkSheet.get_Range(string.Format("{0}2", (char)(99 + i)), string.Format("{0}{1}", (char)(99 + i), 1 + (j + 1) * 5));
                    chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    chartRange.HorizontalAlignment = 3;
                    chartRange.VerticalAlignment = 3;
                }
            }


            // Авторазмер ячеек
            chartRange = xlWorkSheet.UsedRange;
            chartRange.Columns.AutoFit();


            xlWorkBook.SaveAs(string.Format(@"{0}\..\..\sources\LastTimetable.xls", Environment.CurrentDirectory), Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlApp);
            releaseObject(xlWorkBook);
            releaseObject(xlWorkSheet);
                                

        }

        private void useQuery(string query)
        {
            OleDbConnection con = new OleDbConnection(connectionStr);
            con.Open();
            OleDbCommand cmd = new OleDbCommand(query, con);
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            con.Close();
        }

        // --- Event Handlers

        private void buttonGenerate_Click(object sender, EventArgs e)
        {
            DateTime startTime = DateTime.Now;
            buttonGenerate.Enabled = false;
            buttonSave.Enabled = false;
            double[] W = new double[6];

            W[0] = trackBar1.Value / 10.0;
            W[1] = trackBar2.Value / 10.0;
            W[2] = trackBar3.Value / 10.0;
            W[3] = trackBar4.Value / 10.0;
            W[4] = trackBar5.Value / 10.0;
            W[5] = trackBar6.Value / 10.0;

            int magnitudeOfPopulation;
            double mutationProbability;
            if (!Int32.TryParse(textBox1.Text, out magnitudeOfPopulation) || (!Double.TryParse(textBox2.Text, out mutationProbability)) || (magnitudeOfPopulation <= 0) || 
                (mutationProbability > 100) || (mutationProbability <= 0))
            {
                MessageBox.Show("Неверные данные для количества особей или вероятности мутации!");
                buttonGenerate.Enabled = true;
                return;
            }
            mutationProbability = mutationProbability / 100;


            // Создадим популяцию
            double[][] w = new double[magnitudeOfPopulation][];
            for (int i = 0; i < magnitudeOfPopulation; i++)
                w[i] = new double[6];
            Random r = new Random();
            for (int i = 0; i < magnitudeOfPopulation; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    w[i][j] = r.NextDouble();
                }
            }
            // Создадим массив фалгов того, что особь жива.
            bool[] isUnitAlive = new bool[magnitudeOfPopulation];
            for (int i = 0; i < magnitudeOfPopulation; i++) isUnitAlive[i] = true;

            // Запускаем процесс эволюции
            Generator g;
            int indexOfAliveUnit = -1;
            while (true)
            {
                // СКРЕЩИВАНИЕ
                int amountOfAliveUnits = 0;
                // Считаем живых
                for (int i = 0; i < magnitudeOfPopulation; i++)
                {
                    if (isUnitAlive[i]) amountOfAliveUnits++;
                }
                // Сортируем по силе
                for (int i = 0; i < amountOfAliveUnits; i++)
                {
                    double cur_max = 0;
                    int cur_max_index = -1;
                    for (int j = i; j < magnitudeOfPopulation; j++)
                    {
                        if (isUnitAlive[j])
                        {
                            g = new Generator(count, w[j], connectionStr);
                            g.generateTimeTable();
                            double curQuality = g.getMeasureOfQuality(W);
                            if (curQuality > cur_max)
                            {
                                cur_max = curQuality;
                                cur_max_index = j;
                            }
                        }
                    }

                    double[] tmp = w[i];
                    w[i] = w[cur_max_index];
                    w[cur_max_index] = tmp;
                }
                // Скрещиваем попарно и заранее заменяем родителей детьми
                for (int i = 0; i < amountOfAliveUnits / 2; i++)
                {
                    int t = r.Next(5);
                    for (int j = 0; j < 6; j++)
                    {
                        if (j > t)
                        {
                            double tmp = w[2 * i + 1][j];
                            w[2 * i + 1][j] = w[2 * i][j];
                            w[2 * i][j] = tmp;
                        }
                    }
                }



                // МУТАЦИЯ
                int amountOfMutants = 0;
                bool[] didUnitMutate = new bool[magnitudeOfPopulation];
                while (true)
                {
                    if (mutationProbability == 0) break;

                    int mutantIndex = r.Next(magnitudeOfPopulation);
                    
                    // Если особь мертва или уже мутировала - она нам не подходит
                    if (!isUnitAlive[mutantIndex] || didUnitMutate[mutantIndex]) continue;

                    w[mutantIndex][r.Next(6)] = r.NextDouble();

                    didUnitMutate[mutantIndex] = true;
                    amountOfMutants++;
                    if (amountOfMutants >= Math.Ceiling(magnitudeOfPopulation * mutationProbability)) break;
                }



                // СЕЛЕКЦИЯ
                for (int i = 0; i < amountOfAliveUnits / 2; i++)
                {
                    double minQuality = 1000.0;
                    int indexMin = -1;
                    for (int j = 0; j < magnitudeOfPopulation; j++)
                    {
                        if (isUnitAlive[j])
                        {
                            g = new Generator(count, w[j], connectionStr);
                            g.generateTimeTable();
                            double curQuality = g.getMeasureOfQuality(W);
                            if (curQuality <= minQuality)
                            {
                                minQuality = curQuality;
                                indexMin = j;
                            }
                        }
                    }
                    isUnitAlive[indexMin] = false;
                }

                // Если осталась одна особь, заканчиваем цикл
                amountOfAliveUnits = 0;
                for (int i = 0; i < magnitudeOfPopulation; i++)
                {
                    if (isUnitAlive[i])
                    {
                        amountOfAliveUnits++;
                        indexOfAliveUnit = i;
                    }
                }
                if (amountOfAliveUnits == 1) break;
            }



            // Строим окончательное расписание
            g = new Generator(count, w[indexOfAliveUnit], connectionStr);
            g.generateTimeTable();
            gen = g;

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
            buttonGenerate.Enabled = true;
            buttonSave.Enabled = true;
            DateTime endTime = DateTime.Now;
            MessageBox.Show("Время работы: " + (endTime - startTime).ToString());
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы действительно хотите выйти?",
                "Внимание!", MessageBoxButtons.OKCancel) == DialogResult.OK)
                this.Close();
        }

        private void buttonLesson_Click(object sender, EventArgs e)
        {
            AddLesson al = new AddLesson();
            al.Owner = this;
            al.ShowDialog();
            if (!al.query.Equals(null)) useQuery(al.query);
        }

        private void buttonGroup_Click(object sender, EventArgs e)
        {
            AddGroup ag = new AddGroup();
            ag.Owner = this;
            ag.ShowDialog();
            if (!ag.query.Equals(null)) useQuery(ag.query);
        }

        private void buttonRoom_Click(object sender, EventArgs e)
        {
            AddRoom ar = new AddRoom();
            ar.Owner = this;
            ar.ShowDialog();
            if (!ar.query.Equals(null)) useQuery(ar.query);
        }

        private void buttonTeacher_Click(object sender, EventArgs e)
        {
            AddTeacher at = new AddTeacher();
            at.Owner = this;
            at.ShowDialog();
            if (!at.query.Equals(null)) useQuery(at.query);
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            saveInExcel(gen);
            buttonSave.Enabled = false;
        }

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                buttonGenerate_Click(this, new EventArgs());
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                MainForm_KeyDown(this, e);
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                MainForm_KeyDown(this, e);
        }
                        
    }
}
