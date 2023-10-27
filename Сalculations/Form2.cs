using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Сalculations.DataCa;

namespace Сalculations
{
    public partial class Form2 : Form
    {
        private void addRow(DataGridView dtg, Boolean number = false)
        {
            dtg.Rows.Add(1);

            int count = dtg.Rows.Count;
            if (number) { dtg.Rows[count - 1].Cells[0].Value = dtg.Rows.Count; }

        }
        private void addRemove(DataGridView dtg)
        {
            int count = dtg.Rows.Count;
            if (count != 0) { dtg.Rows.RemoveAt(count - 1); }

        }

        private void textFloat_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
        (e.KeyChar != ','))
            {
                e.Handled = true;
            }


            if ((e.KeyChar == ',') && ((sender as TextBox).Text.IndexOf(',') > -1))
            {
                e.Handled = true;
            }
        }



        public DataGridView MyDataGridView
        {
            get { return dataGridView1; }
        }
        public Form2()
        {
            InitializeComponent();
            DataCa.commonMenu = new CommonMenu(this);
            menuStrip1.Items.Add(commonMenu.MenuMenu);

        }
        private void Form2_Load(object sender, EventArgs e)
        {
            if (false)
            {
                string[] name_1 = {
        "Кран мостовий при ТВ = 40%",
        "Зварювальний трансформатор при ТВ = 25% і cosφ = 0,5",
        "Відрізний верстат",
        "Токарно-револьверний верстат",
        "Токарно-гвинторізний верстат",
        "Вертикально-свердлильний верстат",
        "Радіально-свердлильний верстат",
        "Фрезерно-центровий напівавтомат",
        "Горизонтально-фрезерний верстат",
        "Фрезерно-шпонковий верстат",
        "Круглошліфувальний верстат",
        "Вентилятор",
                 "Вентилятор",
                 "Вентилятор",
                 "Вентилятор"};
                int[] count_1 = { 2, 2, 3, 2, 3, 2, 3, 2, 3, 2, 2, 2, 2, 3, 2 };
                double[] Ph = { 55, 15, 22, 9.5, 7.5, 5.5, 11, 9, 10.5, 11, 9.5, 3, 5.6, 5.4, 7.5 };
                double[] Kv = { 0.1, 0.25, 0.16, 0.12, 0.12, 0.12, 0.16, 0.6, 0.16, 0.16, 0.12, 0.6, 0.6, 0.4, 0.5 };
                String[] categories_1 = { "II", "III", "III", "III", "III", "III", "III", "III", "III", "III", "III", "II", "II", "II", "II" };
                int lll = name_1.Length;
                int[] Ttabl1 = { 4, 2, 9, 9, 9, 9, 9, 9, 9, 9, 9, 3, 3, 3, 3 };
                double[] Bk1 = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                double[] Br1 = Enumerable.Repeat(0.67, lll).ToArray();
                double[] Bb1 = Enumerable.Repeat(0.0, lll).ToArray();
                double[] Bo1 = Enumerable.Repeat(0.85, lll).ToArray();
                double[] Bc1 = Enumerable.Repeat(0.0, lll).ToArray();


                int[] Ttabl2 = { 6, 3, 12, 12, 12, 12, 12, 12, 12, 12, 12, 6, 4, 5, 9 };
                double[] Bk2 = Enumerable.Repeat(0.0, lll).ToArray();
                double[] Br2 = Enumerable.Repeat(0.67, lll).ToArray();
                double[] Bb2 = Enumerable.Repeat(0.0, lll).ToArray();
                double[] Bo2 = Enumerable.Repeat(0.7, lll).ToArray();
                double[] Bc2 = Enumerable.Repeat(0.0, lll).ToArray();
                int[] period = Enumerable.Repeat(11, lll).ToArray();


                dataGridView1.Rows.Add(lll);
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.Rows[i].Cells[0].Value = i + 1;
                    dataGridView1.Rows[i].Cells[1].Value = name_1[i];
                    dataGridView1.Rows[i].Cells[2].Value = count_1[i];
                    dataGridView1.Rows[i].Cells[3].Value = Ph[i];
                    dataGridView1.Rows[i].Cells[4].Value = Kv[i];
                    dataGridView1.Rows[i].Cells[5].Value = categories_1[i];
                }
                button1.PerformClick();
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {



                    ///delete
                    ///Ttabl1
                    dataGridView2.Rows[i].Cells[2].Value = Ttabl1[i];
                    dataGridView2.Rows[i].Cells[3].Value = Bk1[i];
                    dataGridView2.Rows[i].Cells[4].Value = Br1[i];
                    dataGridView2.Rows[i].Cells[5].Value = Bb1[i];
                    dataGridView2.Rows[i].Cells[6].Value = Bo1[i];
                    dataGridView2.Rows[i].Cells[7].Value = Bc1[i];

                    dataGridView2.Rows[i].Cells[9].Value = Ttabl2[i];
                    dataGridView2.Rows[i].Cells[10].Value = Bk2[i];
                    dataGridView2.Rows[i].Cells[11].Value = Br2[i];
                    dataGridView2.Rows[i].Cells[12].Value = Bb2[i];
                    dataGridView2.Rows[i].Cells[13].Value = Bo2[i];
                    dataGridView2.Rows[i].Cells[14].Value = Bc2[i];

                    dataGridView2.Rows[i].Cells[16].Value = period[i];
                    ///here
                }
                button2.PerformClick();
                text_repaitPot.Text = "3,2";
                text_repaitKap.Text = "15";
                text_re.Text = "12,5";
                text_coeffNorm.Text = "0,6";
                text_NmR.Text = "0,009";
                text_Kd.Text = "22";
                int asdsa = dataGridView3.Rows.Count;
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    if (true)
                    {
                        List<int> nnnList = Enumerable.Repeat(48, asdsa).ToList();




                        int[] nnn = nnnList.ToArray();
                        List<int> nnnList1 = Enumerable.Repeat(16, asdsa).ToList();




                        int[] nn1n = nnnList1.ToArray();
                        dataGridView3.Rows[i].Cells[3].Value = nn1n[i];
                        dataGridView3.Rows[i].Cells[4].Value = 3;
                        dataGridView3.Rows[i].Cells[7].Value = 11;
                        dataGridView3.Rows[i].Cells[8].Value = "02.19";
                        dataGridView3.Rows[i].Cells[9].Value = "11.21";
                        dataGridView3.Rows[i].Cells[24].Value = 2;
                        dataGridView3.Rows[i].Cells[25].Value = nnn[i];

                        for (int i2 = 0; i2 < lll; i2++)
                        {
                            dataGridView3.Rows[i2].Cells[12].Value = "к";
                            dataGridView3.Rows[i2].Cells[15].Value = "п";
                        }

                    }
                }

            }


            //auto
            if (false)
            {
                string[] name_1 = {
        "Кран мостовий при ТВ = 40%",
        "Зварювальний трансформатор при ТВ = 25% і cosφ = 0,5",
        "Відрізний верстат",
        "Токарно-револьверний верстат",
        "Токарно-гвинторізний верстат",
        "Вертикально-свердлильний верстат",
        "Радіально-свердлильний верстат",
        "Фрезерно-центровий напівавтомат",
        "Горизонтально-фрезерний верстат",
        "Фрезерно-шпонковий верстат",
        "Круглошліфувальний верстат",
        "Вентилятор"};
                int[] count_1 = { 2, 2, 3, 2, 3, 2, 3, 2, 3, 2, 2, 2 };
                double[] Ph = { 55, 15, 22, 9.5, 7.5, 5.5, 11, 9, 10.5, 11, 9.5, 3 };
                double[] Kv = { 0.1, 0.25, 0.16, 0.12, 0.12, 0.12, 0.16, 0.6, 0.16, 0.16, 0.12, 0.6 };
                String[] categories_1 = { "II", "III", "III", "III", "III", "III", "III", "III", "III", "III", "III", "II" };

                int[] Ttabl1 = { 4, 2, 9, 9, 9, 9, 9, 9, 9, 9, 9, 3 };
                double[] Bk1 = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                double[] Br1 = Enumerable.Repeat(0.67, 12).ToArray();
                double[] Bb1 = Enumerable.Repeat(0.0, 12).ToArray();
                double[] Bo1 = Enumerable.Repeat(0.85, 12).ToArray();
                double[] Bc1 = Enumerable.Repeat(0.0, 12).ToArray();


                int[] Ttabl2 = { 6, 3, 12, 12, 12, 12, 12, 12, 12, 12, 12, 6 };
                double[] Bk2 = Enumerable.Repeat(0.0, 12).ToArray();
                double[] Br2 = Enumerable.Repeat(0.67, 12).ToArray();
                double[] Bb2 = Enumerable.Repeat(0.0, 12).ToArray();
                double[] Bo2 = Enumerable.Repeat(0.7, 12).ToArray();
                double[] Bc2 = Enumerable.Repeat(0.0, 12).ToArray();
                int[] period = Enumerable.Repeat(11, 12).ToArray();

                dataGridView1.Rows.Add(12);
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.Rows[i].Cells[0].Value = i + 1;
                    dataGridView1.Rows[i].Cells[1].Value = name_1[i];
                    dataGridView1.Rows[i].Cells[2].Value = count_1[i];
                    dataGridView1.Rows[i].Cells[3].Value = Ph[i];
                    dataGridView1.Rows[i].Cells[4].Value = Kv[i];
                    dataGridView1.Rows[i].Cells[5].Value = categories_1[i];
                }
                button1.PerformClick();
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {



                    ///delete
                    ///Ttabl1
                    dataGridView2.Rows[i].Cells[2].Value = Ttabl1[i];
                    dataGridView2.Rows[i].Cells[3].Value = Bk1[i];
                    dataGridView2.Rows[i].Cells[4].Value = Br1[i];
                    dataGridView2.Rows[i].Cells[5].Value = Bb1[i];
                    dataGridView2.Rows[i].Cells[6].Value = Bo1[i];
                    dataGridView2.Rows[i].Cells[7].Value = Bc1[i];

                    dataGridView2.Rows[i].Cells[9].Value = Ttabl2[i];
                    dataGridView2.Rows[i].Cells[10].Value = Bk2[i];
                    dataGridView2.Rows[i].Cells[11].Value = Br2[i];
                    dataGridView2.Rows[i].Cells[12].Value = Bb2[i];
                    dataGridView2.Rows[i].Cells[13].Value = Bo2[i];
                    dataGridView2.Rows[i].Cells[14].Value = Bc2[i];

                    if (i <= 10)
                    {
                        dataGridView2.Rows[i].Cells[16].Value = period[i];
                    }
                    else
                    {
                        dataGridView2.Rows[i].Cells[16].Value = 7;
                    }
                    ///here
                }
                button2.PerformClick();
                text_repaitPot.Text = "3,2";
                text_repaitKap.Text = "15";
                text_re.Text = "12,5";
                text_coeffNorm.Text = "0,6";
                text_NmR.Text = "0,009";
                text_Kd.Text = "22";


                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    if (true)
                    {
                        List<int> nnnList = Enumerable.Repeat(48, 10).ToList();
                        nnnList.AddRange(Enumerable.Repeat(24, 16));
                        nnnList.AddRange(Enumerable.Repeat(48, 2));

                        List<int> nnnList5 = Enumerable.Repeat(4, 10).ToList();
                        nnnList5.AddRange(Enumerable.Repeat(2, 16));
                        nnnList5.AddRange(Enumerable.Repeat(4, 2));


                        int[] nnn = nnnList.ToArray();
                        List<int> nnnList1 = Enumerable.Repeat(16, 2).ToList();

                        nnnList1.AddRange(Enumerable.Repeat(26, 2));
                        nnnList1.AddRange(Enumerable.Repeat(16, 3));
                        nnnList1.AddRange(Enumerable.Repeat(18, 3));
                        nnnList1.AddRange(Enumerable.Repeat(14, 2));
                        nnnList1.AddRange(Enumerable.Repeat(13, 2));
                        nnnList1.AddRange(Enumerable.Repeat(8, 3));
                        nnnList1.AddRange(Enumerable.Repeat(12, 2));
                        nnnList1.AddRange(Enumerable.Repeat(10, 3));
                        nnnList1.AddRange(Enumerable.Repeat(19, 2));
                        nnnList1.AddRange(Enumerable.Repeat(14, 2));
                        nnnList1.AddRange(Enumerable.Repeat(7, 2));


                        int[] nn1n = nnnList1.ToArray();
                        dataGridView3.Rows[i].Cells[3].Value = nn1n[i];
                        dataGridView3.Rows[i].Cells[4].Value = 3;
                        //dataGridView3.Rows[i].Cells[7].Value = 11;

                        dataGridView3.Rows[i].Cells[8].Value = "02.19";
                        dataGridView3.Rows[i].Cells[9].Value = "11.21";
                        dataGridView3.Rows[i].Cells[24].Value = nnnList5[i];
                        dataGridView3.Rows[i].Cells[25].Value = nnn[i];
                        dataGridView3.Rows[0].Cells[11].Value = "К";
                        dataGridView3.Rows[0].Cells[15].Value = "п";
                        dataGridView3.Rows[0].Cells[19].Value = "п";

                        dataGridView3.Rows[1].Cells[11].Value = "К";
                        dataGridView3.Rows[1].Cells[15].Value = "п";
                        dataGridView3.Rows[1].Cells[19].Value = "п";


                        dataGridView3.Rows[2].Cells[11].Value = "к";
                        dataGridView3.Rows[2].Cells[14].Value = "п";
                        dataGridView3.Rows[2].Cells[17].Value = "п";
                        dataGridView3.Rows[2].Cells[20].Value = "п";

                        dataGridView3.Rows[3].Cells[12].Value = "к";
                        dataGridView3.Rows[3].Cells[15].Value = "п";
                        dataGridView3.Rows[3].Cells[18].Value = "п";
                        dataGridView3.Rows[3].Cells[21].Value = "п";

                        dataGridView3.Rows[4].Cells[12].Value = "к";
                        dataGridView3.Rows[4].Cells[15].Value = "п";

                        dataGridView3.Rows[5].Cells[12].Value = "к";
                        dataGridView3.Rows[5].Cells[15].Value = "п";

                        dataGridView3.Rows[6].Cells[12].Value = "к";
                        dataGridView3.Rows[6].Cells[15].Value = "п";

                        dataGridView3.Rows[7].Cells[12].Value = "к";
                        dataGridView3.Rows[7].Cells[15].Value = "п";

                        dataGridView3.Rows[8].Cells[12].Value = "к";
                        dataGridView3.Rows[8].Cells[15].Value = "п";

                        dataGridView3.Rows[9].Cells[12].Value = "к";
                        dataGridView3.Rows[9].Cells[15].Value = "п";

                        dataGridView3.Rows[10].Cells[12].Value = "к";
                        dataGridView3.Rows[10].Cells[15].Value = "п";

                        dataGridView3.Rows[11].Cells[12].Value = "к";
                        dataGridView3.Rows[11].Cells[15].Value = "п";

                        dataGridView3.Rows[12].Cells[12].Value = "к";
                        dataGridView3.Rows[12].Cells[15].Value = "п";

                        dataGridView3.Rows[13].Cells[12].Value = "к";
                        dataGridView3.Rows[13].Cells[15].Value = "п";
                        for (int i2 = 14; i2 < dataGridView3.Rows.Count - 2; i2++)
                        {
                            dataGridView3.Rows[i2].Cells[12].Value = "к";
                            dataGridView3.Rows[i2].Cells[15].Value = "п";
                        }
                        dataGridView3.Rows[26].Cells[12].Value = "к";
                        dataGridView3.Rows[26].Cells[15].Value = "п";
                        dataGridView3.Rows[26].Cells[16].Value = "п";

                        dataGridView3.Rows[27].Cells[12].Value = "к";
                        dataGridView3.Rows[27].Cells[15].Value = "п";
                        dataGridView3.Rows[27].Cells[16].Value = "п";
                    }
                }
                tabControl1.SelectedTab = yearChart;
                button6.PerformClick();
                button4.PerformClick();
                // this.Hide();
                // DataCa.f4.Show();
            }

        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Diagnostics.Process.GetCurrentProcess().Kill();

        }




        private double MathRound(double number)
        {
            return Math.Round(number, 2);
        }
        private double MathCelling(double number)
        {
            return Math.Round(number, MidpointRounding.AwayFromZero);
        }

        private void btn_addRow_Click(object sender, EventArgs e)
        {
            addRow(dataGridView1, true);
        }

        private void btn_deleteRow_Click(object sender, EventArgs e)
        {
            addRemove(dataGridView1);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Rows.Add(dataGridView1.Rows.Count);

            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                dataGridView2.Rows[i].Cells[0].Value = i + 1;
                dataGridView2.Rows[i].Cells[1].Value = dataGridView1.Rows[i].Cells[1].Value.ToString();
            }

            tabControl1.SelectedTab = repairCycle;
        }
        int countAll;
        private void button2_Click(object sender, EventArgs e)
        {
            countAll = 0;
            dataGridView3.Rows.Clear();
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {

                double coefficient1 = 1, coefficient2 = 1;
                for (int j = 0; j < 5; j++)
                {
                    double cf;
                    double cf1;
                    if (dataGridView2.Rows[i].Cells[3 + j].Value != null)
                    {
                        if (double.TryParse(dataGridView2.Rows[i].Cells[3 + j].Value.ToString(), out cf))
                        {
                            if (cf != 0)
                            {
                                coefficient1 *= cf;
                            }
                        }
                    }

                    if (dataGridView2.Rows[i].Cells[10 + j].Value != null)
                    {
                        if (double.TryParse(dataGridView2.Rows[i].Cells[10 + j].Value.ToString(), out cf1))
                        {
                            if (cf1 != 0)
                            {

                                coefficient2 *= cf1;
                            }
                        }
                    }


                }

                int t_table = int.Parse(dataGridView2.Rows[i].Cells[2].Value.ToString());
                int tpl = (int)Math.Ceiling(MathRound(coefficient1) * t_table);
                dataGridView2.Rows[i].Cells[8].Value = tpl;

                int t_table1 = int.Parse(dataGridView2.Rows[i].Cells[9].Value.ToString());
                int tpl1 = (int)Math.Ceiling(MathRound(coefficient2) * t_table1);
                dataGridView2.Rows[i].Cells[15].Value = tpl1;

                dataGridView2.Rows[i].Cells[16].Value = DataCa.MathRoundInt(12 * tpl / tpl1 - 1);

            }


            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {

                int ccc = int.Parse(dataGridView1.Rows[i].Cells[2].Value.ToString());


                for (int j = 0; j < ccc; j++)
                {

                    int count = dataGridView3.Rows.Count;
                    countAll++;

                    dataGridView3.Rows.Add(1);
                    dataGridView3.Rows[count].Cells[0].Value = count + 1;
                    dataGridView3.Rows[count].Cells[1].Value = dataGridView1.Rows[i].Cells[1].Value;
                    dataGridView3.Rows[count].Cells[2].Value = dataGridView1.Rows[i].Cells[3].Value;
                    dataGridView3.Rows[count].Cells[7].Value = dataGridView2.Rows[i].Cells[16].Value;
                    dataGridView3.Rows[count].Cells[5].Value = dataGridView2.Rows[i].Cells[8].Value;
                    dataGridView3.Rows[count].Cells[6].Value = dataGridView2.Rows[i].Cells[15].Value;



                }
            }


        }


        private void button3_Click(object sender, EventArgs e)
        {

            double repairPot = double.Parse(text_repaitPot.Text);
            double repairKap = double.Parse(text_repaitKap.Text);
            DataCa.tPot = repairPot;
            DataCa.tKap = repairKap;
            double re = double.Parse(text_re.Text);
            double coeffNorm = double.Parse(text_coeffNorm.Text);
            double NmR = double.Parse(text_NmR.Text);
            double Kd = double.Parse(text_Kd.Text);

            double kapSum = 0, potSum = 0, prostSum = 0, remSum = 0;


            DataCa.sumHard = 0;
            DataCa.rAvarage = 0;
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (countAll + 2 != dataGridView3.Rows.Count)
                {
                    dataGridView3.Rows.Add(2);
                }
                if (i < countAll)
                {
                    double Tpot, Tkkap;
                    int zmin = int.Parse(dataGridView3.Rows[i].Cells[4].Value.ToString());
                    double skl = double.Parse(dataGridView3.Rows[i].Cells[3].Value.ToString());

                    int kap = 0, pot = 0;

                    for (int j = 10; j < 22; j++)
                    {
                        if (dataGridView3.Rows[i].Cells[j].Value != null &&
                        dataGridView3.Rows[i].Cells[j].Value.ToString().Equals("К", StringComparison.OrdinalIgnoreCase))
                        {
                            kap += 1;
                        }

                        if (dataGridView3.Rows[i].Cells[j].Value != null &&
                       dataGridView3.Rows[i].Cells[j].Value.ToString().Equals("П", StringComparison.OrdinalIgnoreCase))
                        {
                            pot += 1;
                        }
                    }
                    Tpot = Math.Round((repairPot + repairPot * coeffNorm + NmR * Kd * zmin * re) * pot * skl);
                    Tkkap = Math.Round((repairKap + repairKap * coeffNorm + NmR * Kd * zmin * re) * kap * skl);

                    kapSum += Tkkap;
                    potSum += Tpot;
                    prostSum += int.Parse(dataGridView3.Rows[i].Cells[24].Value.ToString());
                    remSum += int.Parse(dataGridView3.Rows[i].Cells[25].Value.ToString());
                    dataGridView3.Rows[i].Cells[22].Value = Tkkap;
                    dataGridView3.Rows[i].Cells[23].Value = Tpot;
                    DataCa.sumHard += double.Parse(dataGridView3.Rows[i].Cells[3].Value.ToString());
                }
                else
                {
                    dataGridView3.Rows[countAll].Cells[1].Value = "Всього: н-год.";
                    dataGridView3.Rows[countAll + 1].Cells[1].Value = "Загальна трудомісткість ремонтних робіт обладнання дільниці н-год.";
                    dataGridView3.Rows[countAll].Cells[22].Value = kapSum;
                    dataGridView3.Rows[countAll].Cells[23].Value = potSum;
                    dataGridView3.Rows[countAll].Cells[25].Value = remSum;
                    dataGridView3.Rows[countAll].Cells[24].Value = prostSum;
                    dataGridView3.Rows[countAll + 1].Cells[24].Value = kapSum + potSum - remSum - prostSum;
                    DataCa.Trich = int.Parse(dataGridView3.Rows[countAll + 1].Cells[24].Value.ToString());
                }


            }
            DataCa.rAvarage = Math.Round((DataCa.sumHard / countAll * 100)) / 100;
            DataCa.Ta[1] = kapSum - remSum;
            DataCa.Ta[2] = potSum - prostSum;
            DataCa.Ta[0] = DataCa.Ta[1] + DataCa.Ta[2];

        }

        private void yearChart_Enter(object sender, EventArgs e)
        {
            yearChart.ScrollControlIntoView(yearChart);
        }

        private void MenuTool_Click(object sender, EventArgs e)
        {

            this.Hide();

            DataCa.f3.ShowDialog();
        }

        private void EconTool_Click(object sender, EventArgs e)
        {
            this.Hide();
            DataCa.f4.Show();
        }

        private void економічнаЧастинаПродовженняToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            DataCa.f5.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int[] columnWidths = { 30, 280, 30, 35, 30, 30 };
            DataCa.PrintDtg(dataGridView1, "Таблиця 4.1.1 -  Таблиця електронавантаження", columnWidths, 40, true, false);
        }
        private void button5_Click(object sender, EventArgs e)
        {
            int[] columnWidths = { 20, 300, 30, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 30, 30 };
            DataCa.PrintDtg(dataGridView2, "Таблиця 4.1.2 - Розрахунок ремонтного циклу і міжремонтного періоду", columnWidths, 40, true, true);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string text = "РІЧНИЙ ГРАФІК ПЛАНОВО-ПОПЕРЕДЖУВАЛЬНОГО РЕМОНТУ НА 2023 РІК";
            int currentYear = DateTime.Now.Year;

            string updatedText = text.Replace("2023", currentYear.ToString());
            int[] columnWidths = { 30, 150, 40, 30, 30, 30, 30, 40, 40, 40, 25, 25, 25, 25, 25, 25, 25, 25, 25, 25, 25, 25, 50, 50, 60, 30 };
            DataCa.PrintDtg(dataGridView3, text, columnWidths, 40, true, true);
        }
    }
}
