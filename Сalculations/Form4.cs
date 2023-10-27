using DGVPrinterHelper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.ConstrainedExecution;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Сalculations.DataCa;

namespace Сalculations
{
    public partial class Form4 : Form
    {
        private void textFloat_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
        (e.KeyChar != ','))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == ',') && ((sender as TextBox).Text.IndexOf(',') > -1))
            {
                e.Handled = true;
            }
        }
        int Dc, Dv, Dsv, Dpsv, Zm, Ps, Dchv, Ddv, Dlic, Ddo, Fef, Fnom, ChelR, Kzm, Hobsl;

        double Kvn, Fpr, MainPercent, MainSecondaryPercent, AdditionalPercent,
            Fpem, FadditionalL, Fmain, Fadditional, Fannual, AverageSalaryMonthly, Taverage;
        double Fsalary5, Fsalary6, Fprem5, Fprem6, FmainEl5, FmainEl6,
            Fadditional5, FadditionalL5, Fadditional6, FadditionalL6, Fannual5, Fannual6, AverageSalarymounthly5, AverageSalaryMounthly6;
        int FadditionalPercent5, FadditionalPercent6, FmainPercent5, FmainPercent6, FpremPercent5, FpremPercent6, countEl5, countEl6;
        double salaryKap, salaryPot, SalaryPotSurcharge, SalaryKapSurcharge, SalaryMainKap, SalaryMainPot, SalaryPotAdditional, SalaryKapAdditional;
        int PercentSurchargeAdditional, PercentSurchargeMain;
        int Esv, PercentMaterialsKap, PercentMaterialsPot, PercentOperationKapPot, PercentGeneral,
            PercentAdministrative, PercentNonProduction;
        double EsvKap, EsvPot, CostMaterialsPot, CostMaterialsKap, CostsMaintenanceKap, CostsMaintenancePot,
            GeneralCostsPot, GeneralCostsKap, AmdinistrativeCostsKap, AdministrativeCostsPot, NonProduction;


        private double MathRound(double number)
        {
            return Math.Round(number * 100) / 100;
        }
        public Form4()

        {
            InitializeComponent();
            DataCa.commonMenu = new CommonMenu(this);
            menuStrip1.Items.Add(commonMenu.MenuMenu);

        }

        private void Form4_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Diagnostics.Process.GetCurrentProcess().Kill();
        }

        private void menuTool_Click(object sender, EventArgs e)
        {
            this.Hide();
            DataCa.f3.Show();
        }
        private void EconTool_Click(object sender, EventArgs e)
        {
            this.Hide();
            DataCa.f2.Show();
        }
        int checkLoad = 0;
        private void Form4_Load(object sender, EventArgs e)
        {

            if (checkLoad == 0)
            {
                dataGridView1.Rows.Add(4);
                dataGridView1.Rows[0].Cells[0].Value = "Тарифний коефіцієнт";
                dataGridView1.Rows[1].Cells[0].Value = "Погодинна тарифна ставка";
                dataGridView1.Rows[2].Cells[0].Value = "Електрики-ремонтники";
                dataGridView1.Rows[3].Cells[0].Value = "Чергові електрки";
                dataGridView1.Rows[1].Cells[2].Style.BackColor = Color.FromArgb(224, 224, 224);
                dataGridView1.Rows[1].Cells[3].Style.BackColor = Color.FromArgb(224, 224, 224);
                dataGridView1.Rows[1].Cells[4].Style.BackColor = Color.FromArgb(224, 224, 224);
                dataGridView1.Rows[1].Cells[5].Style.BackColor = Color.FromArgb(224, 224, 224);
                dataGridView1.Rows[1].Cells[6].Style.BackColor = Color.FromArgb(224, 224, 224);
                dataGridView1.Rows[1].Cells[2].ReadOnly = true;
                dataGridView1.Rows[1].Cells[3].ReadOnly = true;
                dataGridView1.Rows[1].Cells[4].ReadOnly = true;
                dataGridView1.Rows[1].Cells[5].ReadOnly = true;
                dataGridView1.Rows[1].Cells[6].ReadOnly = true;

                dataGridView1.Rows[3].Cells[1].Style.BackColor = Color.FromArgb(224, 224, 224);
                dataGridView1.Rows[3].Cells[2].Style.BackColor = Color.FromArgb(224, 224, 224);
                dataGridView1.Rows[3].Cells[3].Style.BackColor = Color.FromArgb(224, 224, 224);
                dataGridView1.Rows[3].Cells[4].Style.BackColor = Color.FromArgb(224, 224, 224);
                dataGridView1.Rows[3].Cells[1].ReadOnly = true;
                dataGridView1.Rows[3].Cells[2].ReadOnly = true;
                dataGridView1.Rows[3].Cells[3].ReadOnly = true;
                dataGridView1.Rows[3].Cells[4].ReadOnly = true;
            }
            if (false)
            {
                tDc.Text = "365";
                tDv.Text = "103";
                tDSv.Text = "11";
                tZm.Text = "8";
                tDpsv.Text = "6";
                tPs.Text = "1";
                tDchv.Text = "24";
                tDdv.Text = "4";
                tDlic.Text = "2";
                tDdo.Text = "1";
                tKvn.Text = "1,1";
                //delete
                button1.PerformClick();
                tabControl1.SelectedTab = tabPage2;

                double[] t5_1 = { 1.00, 1.08, 1.23, 1.35, 1.54, 1.8 };
                //   double[] t5_2 = { 28.31, 30.57, 34.82, 38.22, 43.60, 50.96 };
                dataGridView1.Rows[1].Cells[1].Value = 28.31;
                double[] t5_3 = { 0, 0, 0, 4, 4, 2 };
                for (int i = 0; i < dataGridView1.Columns.Count - 1; i++)
                {
                    dataGridView1.Rows[0].Cells[i + 1].Value = t5_1[i];
                    // dataGridView1.Rows[1].Cells[i + 1].Value = t5_2[i];
                    dataGridView1.Rows[2].Cells[i + 1].Value = t5_3[i];
                }
                tKzm.Text = "3";
                tNobsl.Text = "900";
                tKnew.Text = "0,12";
                dataGridView1.Rows[3].Cells[5].Value = 6;
                dataGridView1.Rows[3].Cells[6].Value = 0;


                //delete 
                tMainPercent.Text = "35";
                tSecondaryPercent.Text = "5";
                tAdditionalPercent.Text = "10";
                //
                button2.PerformClick();
                tabControl1.SelectedTab = tabPage3;

                //d

                tFpremPercent5.Text = "35";
                tFmainPercent5.Text = "5";
                tFadditionalPercent5.Text = "10";

                tFpremPercent6.Text = "35";
                tFmainPercent6.Text = "5";
                tFadditionalPercent6.Text = "10";
                button3.PerformClick();

                tPercentSurchargeAdditional.Text = "50";
                tPercentSurchargeMain.Text = "25";
                tEsv.Text = "22";
                tPercentMaterialsKap.Text = "140";
                tPercentMaterialsPot.Text = "120";
                tPercentOperationKap.Text = "300";
                tPercentGeneral.Text = "600";
                tPercentAdministrative.Text = "180";
                tPercentNonProduction.Text = "5";
                tabControl1.SelectedTab = tabPage5;
                button4.PerformClick();
                button6.PerformClick();
                DataCa.f4.Hide();
                DataCa.f5.Show();

            }


        }
        private void button1_Click(object sender, EventArgs e)
        {

            Dc = int.Parse(tDc.Text);
            Dv = int.Parse(tDv.Text);
            Dsv = int.Parse(tDSv.Text);
            Dpsv = int.Parse(tDpsv.Text);
            Zm = int.Parse(tZm.Text);
            Ps = int.Parse(tPs.Text);
            Dchv = int.Parse(tDchv.Text);
            Ddv = int.Parse(tDdv.Text);
            Dlic = int.Parse(tDlic.Text);
            Ddo = int.Parse(tDdo.Text);
            Kvn = double.Parse(tKvn.Text);
            DataCa.daysCalendar = Dc;
            DataCa.dayHoliday = Dsv;
            DataCa.dayPreHoliday = Dpsv;
            DataCa.shiftDuration = Zm;
            DataCa.dayWekkend = Dv;
            Fnom = (Dc - Dv - Dsv) * Zm - Dpsv * Ps;
            Fef = Fnom - (Dchv + Ddv + Dlic + Ddo) * Zm;
            double ss = DataCa.Trich / (Fef * Kvn);
            ChelR = (int)Math.Ceiling(Math.Round(ss, 1));

            label2.Text = "Річний дійсний (ефективний) фонд часу роби робітника Феф - " + Fef.ToString();
            label3.Text = "Номінальний річний фонд робочого часу одного робітника  Фном - " + Fnom.ToString();
            label4.Text = "Облікова чисельність електриків-ремонтників, що входять до складу бригади Чел.р., чол. - " + ChelR.ToString();


        }



        private void button2_Click(object sender, EventArgs e)
        {

            Kzm = int.Parse(tKzm.Text);
            Hobsl = int.Parse(tNobsl.Text);
            double Knew = double.Parse(tKnew.Text);
            int Drob = Dc - Dv - Dsv;
            int DGraph = (int)Math.Round(((Dc - Dv - Dsv - Dchv) * (1 - Knew)));
            double Ksp = Math.Ceiling((double)Drob / DGraph * 100) / 100; ;
            double Chyav = Math.Round((double)(DataCa.sumHard * Kzm) / Hobsl, 2);
            double Cho = Math.Ceiling((double)Chyav * Ksp * 100) / 100; ;

            for (int i = 2; i < dataGridView1.Columns.Count; i++)
            {
                double cef = double.Parse(dataGridView1.Rows[0].Cells[i].Value.ToString());
                double tarif = double.Parse(dataGridView1.Rows[1].Cells[1].Value.ToString());
                dataGridView1.Rows[1].Cells[i].Value = Math.Round(tarif * cef * 100) / 100;

            }

            double max = 0;
            double min = 99999999;
            int allPerson = 0;
            int average = 0;
            int mindischarge = 0;
            Boolean first = true;
            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {

                Boolean isPerson = double.Parse(dataGridView1.Rows[2].Cells[i].Value.ToString()) > 0;


                int person = int.Parse(dataGridView1.Rows[2].Cells[i].Value.ToString());
                double maxS = double.Parse(dataGridView1.Rows[1].Cells[i].Value.ToString());
                double minS = double.Parse(dataGridView1.Rows[1].Cells[i].Value.ToString());
                int discharge = i;
                if (first && isPerson)
                {
                    mindischarge = i;
                    first = !first;
                }
                if (maxS > max && isPerson)
                {
                    max = double.Parse(dataGridView1.Rows[1].Cells[i].Value.ToString());
                }

                if (isPerson)
                {
                    allPerson += person;
                    average += person * discharge;
                }
            }

            countEl5 = int.Parse(dataGridView1.Rows[3].Cells[5].Value.ToString());

            countEl6 = int.Parse(dataGridView1.Rows[3].Cells[6].Value.ToString());

            double Paverage = Math.Round((double)average / allPerson * 100) / 100;
            int PaverageInt = (int)Math.Floor((double)average / allPerson);

            min = double.Parse(dataGridView1.Rows[1].Cells[PaverageInt].Value.ToString());

            Taverage = Math.Round(((max - min) * (Paverage - PaverageInt) + min) * 100) / 100;


            label6.Text = "Середній розряд робочих Рсер - " + Paverage;
            label7.Text = "Середня годинна тарифна ставку Тст.сер - " + Taverage;
            label11.Text = "Явочну чисельність чергових електриків Чяв, чол. - " + Chyav.ToString() + ", згідно нормативно-правової бази, явочна чисельність чергових електриків має складати по дві особи на зміну.";
            label13.Text = "Облікову чисельність чергових електриків скоригуємо коефіцієнтом спискового складу Чо , чол. - " + Cho.ToString();
            label14.Text = "Коефіцієнт спискового складу Ксп, чол. - " + Ksp.ToString();
            label15.Text = "Робочі дні за відповідний рік Дроб - " + Drob.ToString();
            label16.Text = "Кількість днів виходу одного робітника за графіком Дграф - " + DGraph.ToString();
            label17.Text = "Коефіцієнт, що враховує втрати робочого часу з поважних причин Квих - " + (1 - Knew).ToString();


            Fpr = Math.Round(Taverage * DataCa.Trich * 100) / 100;


        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            MainPercent = double.Parse(tMainPercent.Text);
            MainSecondaryPercent = double.Parse(tSecondaryPercent.Text);
            AdditionalPercent = double.Parse(tAdditionalPercent.Text);

            Fpem = DataCa.MathRound(MainPercent * Fpr / 100);
            FadditionalL = DataCa.MathRound(MainSecondaryPercent * Fpr / 100);
            Fmain = DataCa.MathRound(Fpem + FadditionalL + Fpr);
          // MessageBox.Show(Fpem.ToString());
          //  MessageBox.Show(FadditionalL.ToString());
           // MessageBox.Show(Fpr.ToString());
            Fadditional = DataCa.MathRound((AdditionalPercent * Fmain) / 100);
            Fannual = DataCa.MathRound(Fmain + Fadditional);
            AverageSalaryMonthly = DataCa.MathRound(Fannual / (ChelR * 12));


            if (countEl5 > 0)
            {

                FpremPercent5 = int.Parse(tFpremPercent5.Text);
                FmainPercent5 = int.Parse(tFmainPercent5.Text);
                FadditionalPercent5 = int.Parse(tFadditionalPercent5.Text);

                Fsalary5 = MathRound((double)dataGridView1.Rows[1].Cells[5].Value * Fef * countEl5);
                Fprem5 = MathRound((FpremPercent5 * Fsalary5) / 100);
                FadditionalL5 = MathRound((FmainPercent5 * Fsalary5) / 100);
                FmainEl5 = MathRound(Fsalary5 + Fprem5 + FadditionalL5);

                Fadditional5 = MathRound((FadditionalPercent5 * FmainEl5) / 100);
                Fannual5 = MathRound(FmainEl5 + Fadditional5);

            }
            else
            {
                Fsalary5 = 0;
                Fprem5 = 0;
                FadditionalL5 = 0;
                FmainEl5 = 0;

                Fadditional5 = 0;
                Fannual5 = 0;
            }
            if (countEl6 > 0)
            {

                FpremPercent6 = int.Parse(tFpremPercent6.Text);
                FmainPercent6 = int.Parse(tFmainPercent6.Text);
                FadditionalPercent6 = int.Parse(tFadditionalPercent6.Text);

                Fsalary6 = MathRound((double)dataGridView1.Rows[1].Cells[6].Value * Fef * countEl6);
                Fprem6 = MathRound((FpremPercent6 * Fsalary6) / 100);
                FadditionalL6 = MathRound((FmainPercent6 * Fsalary6) / 100);
                FmainEl6 = MathRound(Fsalary6 + Fprem6 + FadditionalL6);


                Fadditional6 = MathRound((FadditionalPercent6 * FmainEl6) / 100);
                Fannual6 = MathRound(FmainEl6 + Fadditional6);
            }
            else
            {
                Fsalary6 = 0;
                Fprem6 = 0;
                FadditionalL6 = 0;
                FmainEl6 = 0;

                Fadditional6 = 0;
                Fannual6 = 0;
            }
            double Fall = Fannual5 + Fannual6;
            dataGridView2.Rows.Add(4);
            dataGridView2.Rows[0].Cells[0].Value = "Електрики-ремонтники";
            dataGridView2.Rows[1].Cells[0].Value = "Чергові електрики";
            dataGridView2.Rows[2].Cells[0].Value = "Чергові електрики";
            dataGridView2.Rows[3].Cells[0].Value = "Всього";

            dataGridView2.Rows[0].Cells[4].Value = "-";
            dataGridView2.Rows[1].Cells[4].Value = "V";
            dataGridView2.Rows[2].Cells[4].Value = "VI";

            if (countEl5 > 0)
            {

                AverageSalarymounthly5 = MathRound(Fall / (countEl5 * 12));
            }
            else
            {
                AverageSalarymounthly5 = 0;
            }
            if (countEl6 > 0)
            {

                AverageSalaryMounthly6 = MathRound(Fall / (countEl6 * 12));

            }
            else
            {
                AverageSalaryMounthly6 = 0;
            }
            dataGridView2.Rows[0].Cells[1].Value = DataCa.Trich;
            dataGridView2.Rows[1].Cells[1].Value = Fef;
            dataGridView2.Rows[2].Cells[1].Value = Fef;

            dataGridView2.Rows[0].Cells[2].Value = MathRound(Taverage);
            dataGridView2.Rows[1].Cells[2].Value = MathRound(double.Parse(dataGridView1.Rows[1].Cells[5].Value.ToString()));
            dataGridView2.Rows[2].Cells[2].Value = MathRound(double.Parse(dataGridView1.Rows[1].Cells[6].Value.ToString()));

            dataGridView2.Rows[0].Cells[3].Value = ChelR;
            dataGridView2.Rows[1].Cells[3].Value = countEl5;
            dataGridView2.Rows[2].Cells[3].Value = countEl6;

            dataGridView2.Rows[0].Cells[5].Value = 0;
            dataGridView2.Rows[1].Cells[5].Value = Fef;
            dataGridView2.Rows[2].Cells[5].Value = Fef;

            dataGridView2.Rows[0].Cells[6].Value = Fmain;
            dataGridView2.Rows[1].Cells[6].Value = FmainEl5;
            dataGridView2.Rows[2].Cells[6].Value = FmainEl6;

            dataGridView2.Rows[0].Cells[7].Value = AdditionalPercent;
            dataGridView2.Rows[1].Cells[7].Value = FadditionalPercent5;
            dataGridView2.Rows[2].Cells[7].Value = FadditionalPercent6;

            dataGridView2.Rows[0].Cells[8].Value = Fadditional;
            dataGridView2.Rows[1].Cells[8].Value = Fadditional5;
            dataGridView2.Rows[2].Cells[8].Value = Fadditional6;

            dataGridView2.Rows[0].Cells[9].Value = Fannual;
            dataGridView2.Rows[1].Cells[9].Value = Fannual5;
            dataGridView2.Rows[2].Cells[9].Value = Fannual6;

            dataGridView2.Rows[0].Cells[10].Value = AverageSalaryMonthly;
            dataGridView2.Rows[1].Cells[10].Value = AverageSalarymounthly5;
            dataGridView2.Rows[2].Cells[10].Value = AverageSalaryMounthly6;

            //row - down/ cell - right

            int personAll = 0;
            double FondAll = 0, FondAdditional = 0, FondAnnual = 0;
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {

                if (int.TryParse(dataGridView2.Rows[i].Cells[3].Value.ToString(), out int parsedValue))
                {

                    personAll += parsedValue;
                }
                if (double.TryParse(dataGridView2.Rows[i].Cells[6].Value.ToString(), out double parsedValue2))
                {

                    FondAll += parsedValue2;
                }
                if (double.TryParse(dataGridView2.Rows[i].Cells[8].Value.ToString(), out double parsedValue3))
                {

                    FondAdditional += parsedValue3;
                }
                if (double.TryParse(dataGridView2.Rows[i].Cells[9].Value.ToString(), out double parsedValue4))
                {

                    FondAnnual += parsedValue4;
                }
            }
            FondAll = DataCa.MathRound(FondAll);
            FondAdditional = DataCa.MathRound(FondAdditional);
            FondAnnual = DataCa.MathRound(FondAnnual);
            dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[3].Value = personAll;
            dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[6].Value = FondAll;
            dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[8].Value = FondAdditional;
            dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[9].Value = FondAnnual;
            tabControl1.SelectedTab = tabPage4;



            DataCa.Ta[7] = double.Parse(dataGridView2.Rows[0].Cells[3].Value.ToString());
            DataCa.Ta[8] = double.Parse(dataGridView2.Rows[1].Cells[3].Value.ToString()) + double.Parse(dataGridView2.Rows[2].Cells[3].Value.ToString());

            DataCa.Ta[9] = Fannual;
            DataCa.Ta[10] = DataCa.MathRound(Fannual5 + Fannual6);

            DataCa.Ta[11] = AverageSalaryMonthly;
            DataCa.Ta[12] = DataCa.MathRound(AverageSalarymounthly5 + AverageSalaryMounthly6);

            DataCa.Ta[18] = DataCa.MathRound(FondAnnual / personAll);
            DataCa.Ta[19] = DataCa.MathRound(DataCa.Trich / double.Parse(dataGridView2.Rows[0].Cells[3].Value.ToString()));

        }

        private void bPrintTable5_3_1_Click(object sender, EventArgs e)
        {
            int[] columnWidths = { 80, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40 };
            DataCa.PrintDtg(dataGridView2, "Таблиця 5.3.1 - Розрахунок річного фонду заробітної плати ремонтного персоналу цеха  та середньомісячної зарплати робітника", columnWidths, 130, true, true);

        }
        private void InputNumber(object sender, KeyPressEventArgs e)
        {
            DataCa.TextFloat_KeyPress(sender, e);
        }
        private void button4_Click(object sender, EventArgs e)
        {
            PercentSurchargeAdditional = int.Parse(tPercentSurchargeAdditional.Text);
            PercentSurchargeMain = int.Parse(tPercentSurchargeMain.Text);
            Esv = int.Parse(tEsv.Text);
            PercentMaterialsKap = int.Parse(tPercentMaterialsKap.Text);
            PercentMaterialsPot = int.Parse(tPercentMaterialsPot.Text);
            PercentOperationKapPot = int.Parse(tPercentOperationKap.Text);
            PercentGeneral = int.Parse(tPercentGeneral.Text);
            PercentAdministrative = int.Parse(tPercentAdministrative.Text);
            PercentNonProduction = int.Parse(tPercentNonProduction.Text);



            salaryKap = DataCa.MathRound(DataCa.tKap * Taverage * DataCa.rAvarage);
            salaryPot = DataCa.MathRound(DataCa.tPot * Taverage * DataCa.rAvarage);

            SalaryKapSurcharge = DataCa.MathRound((PercentSurchargeMain * salaryKap) / 100);
            SalaryPotSurcharge = DataCa.MathRound((PercentSurchargeMain * salaryPot) / 100);

            SalaryMainKap = DataCa.MathRound(SalaryKapSurcharge + salaryKap);
            SalaryMainPot = DataCa.MathRound(SalaryPotSurcharge + salaryPot);

            SalaryKapAdditional = DataCa.MathRound((SalaryMainKap * PercentSurchargeAdditional) / 100);
            SalaryPotAdditional = DataCa.MathRound((SalaryMainPot * PercentSurchargeAdditional) / 100);
            /*  
              double EsvKap, EsvPot, CostMaterialsPot, CostMaterialsKap, CostsMaintenanceKap, CostsMaintenancePot,
                  GeneralCostsPot, GeneralCostsKap, AmdinistrativeCostsKap, AdministrativeCostsPot, NonProduction;*/
            EsvKap = DataCa.MathRound(Esv * (SalaryMainKap + SalaryKapAdditional) / 100);
            EsvPot = DataCa.MathRound(Esv * (SalaryMainPot + SalaryPotAdditional) / 100);

            CostMaterialsKap = DataCa.MathRound((PercentMaterialsKap * SalaryMainKap) / 100);
            CostMaterialsPot = DataCa.MathRound((PercentMaterialsPot * SalaryMainPot) / 100);

            CostsMaintenanceKap = DataCa.MathRound((PercentOperationKapPot * SalaryMainKap) / 100);
            CostsMaintenancePot = DataCa.MathRound((PercentOperationKapPot * SalaryMainPot) / 100);

            GeneralCostsKap = DataCa.MathRound((PercentGeneral * SalaryMainKap) / 100);
            GeneralCostsPot = DataCa.MathRound((PercentGeneral * SalaryMainPot) / 100);

            AmdinistrativeCostsKap = DataCa.MathRound((PercentAdministrative * SalaryMainKap) / 100);
            AdministrativeCostsPot = DataCa.MathRound((PercentAdministrative * SalaryMainPot) / 100);

            NonProduction =
                DataCa.MathRound(
                (PercentNonProduction * (SalaryMainKap + SalaryKapAdditional + CostMaterialsKap + EsvKap + CostsMaintenanceKap + GeneralCostsKap + AmdinistrativeCostsKap)
                ) / 100);


            label18.Text = "Пряма заробітна плата електриків-ремонтників - " + salaryPot.ToString();
            label19.Text = "Пряма заробітна плата електриків-ремонтників - " + salaryKap.ToString();

            label20.Text = "Доплата - " + SalaryKapSurcharge.ToString();
            label21.Text = "Доплата - " + SalaryPotSurcharge.ToString();


            label22.Text = "Фонд основної заробітної плати - " + SalaryMainKap.ToString();
            label23.Text = "Фонд основної заробітної плати - " + SalaryMainPot.ToString();

            label24.Text = "Додаткова заробітна плата - " + SalaryKapAdditional.ToString();
            label25.Text = "Додаткова заробітна плата - " + SalaryPotAdditional.ToString();

            label26.Text = "ЄСВ - " + EsvPot.ToString();
            label27.Text = "Вартість матеріалів для виконання ремонтних робіт - " + CostMaterialsPot.ToString();
            label28.Text = "Витрати на утримання та експлуатацію обладнання - " + CostsMaintenancePot.ToString();
            label29.Text = "Загальновиробничі витрати – " + GeneralCostsPot.ToString();
            label30.Text = "Адміністративні витрати - " + AdministrativeCostsPot.ToString();

            label31.Text = "ЄСВ - " + EsvKap.ToString();
            label32.Text = "Вартість матеріалів для виконання ремонтних робіт - " + CostMaterialsKap.ToString();
            label33.Text = "Витрати на утримання та експлуатацію обладнання - " + CostsMaintenanceKap.ToString();
            label34.Text = "Загальновиробничі витрати – " + GeneralCostsKap.ToString();
            label35.Text = "Адміністративні витрати - " + AmdinistrativeCostsKap.ToString();
            label36.Text = "Позавиробничі витрати складають -  " + NonProduction.ToString();

            string[] articleNames = new string[]
            {
                "Матеріальні витрати",
                "Основана заробітна плата",
                "Додаткова заробітна плата",
                "Відрахування на єдиний соціальний внесок",
                "Витрати на утримання та експлуатацію обладнання",
                "Загальновиробничі витрати",
                "Виробнича собівартість",
                "Адміністративні витрати",
                "Позавиробничі витрати",
                "Повна собівартість"
                        };
            dataGridView3.Rows.Clear();
            foreach (string name in articleNames)
            {
                dataGridView3.Rows.Add(name);
            }
            dataGridView3.Rows[0].Cells[1].Value = CostMaterialsPot.ToString();
            dataGridView3.Rows[0].Cells[3].Value = CostMaterialsKap.ToString();

            dataGridView3.Rows[1].Cells[1].Value = SalaryMainPot.ToString();
            dataGridView3.Rows[1].Cells[3].Value = SalaryMainKap.ToString();

            dataGridView3.Rows[2].Cells[1].Value = SalaryPotAdditional.ToString();
            dataGridView3.Rows[2].Cells[3].Value = SalaryKapAdditional.ToString();

            dataGridView3.Rows[3].Cells[1].Value = EsvPot.ToString();
            dataGridView3.Rows[3].Cells[3].Value = EsvKap.ToString();

            dataGridView3.Rows[4].Cells[1].Value = CostsMaintenancePot.ToString();
            dataGridView3.Rows[4].Cells[3].Value = CostsMaintenanceKap.ToString();

            dataGridView3.Rows[5].Cells[1].Value = GeneralCostsPot.ToString();
            dataGridView3.Rows[5].Cells[3].Value = GeneralCostsKap.ToString();

            double sum1 = 0, sum2 = 0;
            for (int i = 0; i < 6; i++)
            {
                sum1 += MathRound(double.Parse(dataGridView3.Rows[i].Cells[1].Value.ToString()));
                sum2 += MathRound(double.Parse(dataGridView3.Rows[i].Cells[3].Value.ToString()));
            }
            dataGridView3.Rows[6].Cells[1].Value = MathRound(sum1).ToString();
            dataGridView3.Rows[6].Cells[3].Value = MathRound(sum2).ToString();

            dataGridView3.Rows[7].Cells[1].Value = AdministrativeCostsPot.ToString();
            dataGridView3.Rows[7].Cells[3].Value = AmdinistrativeCostsKap.ToString();

            dataGridView3.Rows[8].Cells[3].Value = NonProduction.ToString();

            dataGridView3.Rows[9].Cells[1].Value = MathRound(sum1 + AdministrativeCostsPot);
            dataGridView3.Rows[9].Cells[3].Value = MathRound(sum2 + NonProduction + AmdinistrativeCostsKap);
            double max1 = MathRound(sum1 + AdministrativeCostsPot);
            double max2 = MathRound(sum2 + NonProduction + AmdinistrativeCostsKap);

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {


                if (i != 6 && i != 8)
                {
                    double current1 = double.Parse(dataGridView3.Rows[i].Cells[1].Value.ToString());
                    double procent1 = MathRound(current1 / max1 * 100);
                    dataGridView3.Rows[i].Cells[2].Value = procent1;
                }
                if (i != 6)
                {
                    double current2 = double.Parse(dataGridView3.Rows[i].Cells[3].Value.ToString());
                    double procent2 = MathRound(current2 / max2 * 100);
                    dataGridView3.Rows[i].Cells[4].Value = procent2;

                }


            }

            DataCa.Ta[5] = MathRound(double.Parse(dataGridView3.Rows[dataGridView3.RowCount - 1].Cells[1].Value.ToString()));
            DataCa.Ta[4] = MathRound(double.Parse(dataGridView3.Rows[dataGridView3.RowCount - 1].Cells[3].Value.ToString()));
            DataCa.Ta[3] = MathRound(DataCa.Ta[4] + DataCa.Ta[5]);


        }

        private void button5_Click(object sender, EventArgs e)
        {
            int[] columnWidths = { 80, 40, 40, 40, 40 };
            DataCa.PrintDtg(dataGridView3, "Таблиця 5.4.1 – Калькуляція собівартості планових ремонтів", columnWidths, 60, true, false);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Hide();
            DataCa.f5.Show();
        }

        private void економічнаЧастинаПродовженняToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            DataCa.f5.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string text = "Таблиця 5.1.1 Погодинні тарифні ставки електриків-ремонтників, чергових електриків";

            int[] columnWidths = { 200, 40, 40, 40, 40, 40, 40 };
            DataCa.PrintDtg(dataGridView1, text, columnWidths, 40, true, false);
        }
    }
}
