using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static Сalculations.DataCa;

namespace Сalculations
{
    public partial class Form5 : Form
    {
        int PercentSalary, PercentEsv, PercentAdministrative, PercentGeneral, PercentSum, PercentAdditional;
        double PremiumSlary, MainSalary, SalaryAdditional, Esv, GeneralSalary,
            AdministrativeSalaty, FullCost, SumProfit, CostsWorks;
        double CoefficientRepair, CoefficientLight, CountShift, CountHourLight,
        Power, ProcentLosses, CountLightPoint, CoefficientLoad, TariffLight, AreaMachine, CoefficientAreaMain, CostElectricity;
        int AreaGeneral, AreaMain;
        double MainSalaryByElectricity;
        double profitAll, AllCosts;
        public Form5()
        {
            InitializeComponent();
            DataCa.commonMenu = new CommonMenu(this);
            menuStrip1.Items.Add(commonMenu.MenuMenu);

            string[,] data3 = new string[,]
               {
                  {"Витрати на придбання електрообладнання" ,"0"},
                  {"Транспортно-заготівельні витрати " ,"0"},
                  {"Всього  " ,"0"},
                  {"Основна заробітна плата " ,"0"},
                  {"Додаткова зарплата " ,"0"},
                  {"Відрахування на єдиний соціальний внесок " ,"0"},
                  {"Загальновиробничі витрати " ,"0"},
                  {"Адміністративні витрати " ,"0"},
                  {"Всього витрат ( Повна собівартість) " ,"0"},
                  {"Нормативний прибуток " ,"0"},
                  {"Вартість робіт за внутрішніми розрахунковими цінами " ,"0"},

               };


            for (int i = 0; i < data3.GetLength(0); i++)
            {
                dataGridView3.Rows.Add(data3[i, 0], data3[i, 1]);
            }



        }




        private void Form5_Load(object sender, EventArgs e)
        {



            if (false)
            {
                string[] data = new string[]
{
                "Трансформатор ТБС3-0,16", "шт", "12", "200,00",
                "ЕлектромагнітМИСА4100ЕУ3", "шт", "6", "140,00",
                "Пускач реверсивний ПМЛ210004, Ін=25А", "шт", "14", "25,00",
                "Пускач реверсивний  ПМЛ110004, Ін=10 А", "шт", "7", "16,00",
                "Пускач ПМЛ1 10004, Ін=10А", "шт", "6", "22,00",
                "Реле теплове РТЛ205304, Ін=80А, Іте=27,5А", "шт", "5", "88,00",
                "Реле теплове РТЛ100304, Ін=0,31 А", "шт", "5", "95,00",
                "Реле теплове РТЛ101404, Ін=25А, Іте=8,5А", "шт", "6", "85,00",
                "Запобіжник ПР-2, Ін=100А, Івст=80А", "шт", "25", "55,00",
                "Запобіжник ПР-2, Ін=60А, Івст=20А", "шт", "10", "69,00",
                "Запобіжник ПР-2, Ін=15А, Івст=6А", "шт", "10", "45,00",
                "Запобіжник ПР-2, Ін=15А, Івст=6А", "шт", "8", "44,00",
                "Вимикач пакетний ПВ2-60, Ін=60А", "шт", "8", "320,00",
                "Перемикач пакетний ПП1-60, Ін=40А", "шт", "8", "450,00",
                "Вимикач пакетний ПКП-25, Ін=29А", "шт", "6", "600,00",
                "Перемикач ПКП25-2-57У2", "шт", "6", "780,00",
                "Командо апарат КА4048-2У2", "шт", "1", "1000,00",
                "Вимикач освітлення ПУ-2М", "шт", "2", "250,00",
                "Кнопка КЕ-021 червона", "шт", "20", "55,00",
                "Кнопка КЕ-011 чорна", "шт", "20", "55,00",
                "Кінцевий вимикач ВПК2111АУ2", "шт", "6", "150,00",
                "Реле контролю швидкості РКН028", "шт", "3", "560,00",
                "Резистори МЛТ 10кОм+10%", "шт", "25", "45,00",
                "Вимикач АЕ2956М-100-20; 660В", "шт", "25", "200,00",
                "Вимикач АЕ2056М-100-20; 110В", "шт", "4", "150,00",
                "Реле електротеплове струмове РТЛ1021", "шт", "14", "500,00",
                "Реле електротеплове струмове РТЛ2053", "шт", "13", "450,00",
                "Реле електротеплове струмове РТЛ2057", "шт", "12", "460,00",
                "Реле електротеплове струмове РТЛ1004", "шт", "12", "380,00",
                "Запобіжник ПРС-6-П з плавкою вставкою ПВД1-4", "шт", "28", "15,00",
                "Пускач магнітний", "шт", "25", "80,00",
                "П6-111,10В", "шт", "16", "108,00",
                "Пускач магнітний", "шт", "17", "123,00",
                "ПМЛ 3100,110В", "шт", "7", "100,00",
                "Пускач магнітний", "шт", "15", "188,00",
                "ПМЛ 4100,110В", "шт", "8", "98,00",
                "Пускач магнітний", "шт", "2", "45000,00",
                "ПМЛ 4100,110В", "шт", "40", "15,00",
                "Електродвигун 4А132S4 7,5кВт 1500 об/хв", "шт", "35", "15,00",
                "Резистор ПЄВ-12 120 Ом", "шт", "55", "15,00",
                "Діод Д243", "шт", "11", "155,00",
                "Діод Д2226В", "шт", "16", "180,00",
                "Перемикач ПКУЗ", "шт", "8", "134,00",
                "Перемикач 11СМ2015", "шт", "9", "144,00",
                "Перемикач ПЕ-061", "шт", "7", "166,00",
                "Перемикач ПГК-11П 2Н-8А", "шт", "3", "138,00",
                "Вимикач", "шт", "8", "145,00"
         };

                for (int i = 0; i < data.Length; i += 4)
                {
                    dataGridView1.Rows.Add(data.Skip(i).Take(4).ToArray());

                }
                textBox1.Text = "15";
                button11.PerformClick();



                tabControl1.SelectedTab = tabPage2;


                string[,] data2 = new string[,]
                {
                    {"Трансформатор ТБС3-0,16", "250,00", "10"},
                    {"ЕлектромагнітМИСА4100ЕУ3", "120,00", "22"},
                    {"Пускач реверсивний ПМЛ210004, Ін=25А", "20,00", "10"},
                    {"Пускач реверсивний ПМЛ110004, Ін=10 А", "25,00", "5"},
                    {"Пускач ПМЛ1 10004, Ін=10А", "38,00", "4"},
                    {"Реле теплове РТЛ205304, Ін=80А, Іте=27,5А", "45,00", "5"},
                    {"Реле теплове РТЛ100304, Ін=0,31 А", "55,00", "5"},
                    {"Реле теплове РТЛ101404, Ін=25А, Іте=8,5А", "55,00", "3"},
                    {"Запобіжник ПР-2, Ін=100А, Івст=80А", "3,50", "20"},
                    {"Запобіжник ПР-2, Ін=60А, Івст=20А", "3,50", "10"},
                    {"Запобіжник ПР-2, Ін=15А, Івст=6А", "3,50", "10"},
                    {"Запобіжник ПР-2, Ін=15А, Івст=6А", "3,50", "20"},
                    {"Вимикач пакетний ПВ2-60, Ін=60А", "60,00", "4"},
                    {"Перемикач пакетний ПП1-60, Ін=40А", "33,00", "5"},
                    {"Вимикач пакетний ПКП-25, Ін=29А", "28,00", "6"},
                    {"Перемикач ПКП25-2-57У2", "25,00", "15"},
                    {"Командо апарат КА4048-2У2", "450,00", "10"},
                    {"Вимикач освітлення ПУ-2М", "45,00", "2"},
                    {"Кнопка КЕ-021 червона", "12,00", "10"},
                    {"Кнопка КЕ-011 чорна", "12,00", "10"},
                    {"Кінцевий вимикач ВПК2111АУ2", "36,00", "3"},
                    {"Реле контролю швидкості РКН028", "22,00", "3"},
                    {"Резистори МЛТ 10кОм+10%", "3,50", "10"},
                    {"Вимикач АЕ2956М-100-20; 660В", "24,00", "10"},
                    {"Вимикач АЕ2056М-100-20; 110В", "33,00", "1"},
                    {"Реле електротеплове струмове РТЛ1021", "20,00", "2"},
                    {"Реле електротеплове струмове РТЛ2053", "20,00", "2"},
                    {"Реле електротеплове струмове РТЛ2057", "20,00", "2"},
                    {"Реле електротеплове струмове РТЛ1004", "20,00", "20"},
                    {"Запобіжник ПРС-6-П з плавкою вставкою ПВД1-4", "3,50", "20"},
                    {"Пускач магнітний", "55,00", "20"},
                    {"П6-111,10В", "45,00", "4"},
                    {"Пускач магнітний", "45,00", "10"},
                    {"ПМЛ 3100,110В", "30,00", "20"},
                    {"Пускач магнітний", "45,00", "10"},
                    {"ПМЛ 4100,110В", "30,00", "10"},
                    {"Пускач магнітний", "45,00", "10"},
                    {"ПМЛ 4100,110В", "30,00", "10"},
                    {"Електродвигун 4А132S4 7,5кВт 1500 об/хв", "10,00", "30"},
                    {"Резистор ПЄВ-12 120 Ом", "4,50", "20"},
                    {"Діод Д243", "6,00", "11"},
                    {"Діод Д2226В", "6,00", "10"},
                    {"Перемикач ПКУЗ", "32,00", "10"}
                };



                // Добавьте строки из массива в DataGridView
                for (int i = 0; i < data2.GetLength(0); i++)
                {
                    dataGridView2.Rows.Add(data2[i, 0], data2[i, 1], data2[i, 2]);
                }

                button6.PerformClick();
                tPercentSalary.Text = "40";
                tPercentAdditional.Text = "10";
                tPercentEsv.Text = "22";
                tPercentGeneral.Text = "250";
                tPercentAdministrative.Text = "190";
                tPercentSum.Text = "40";

                tabControl1.SelectedTab = tabPage3;
                button7.PerformClick();
                tCountMachine.Text = "28";
                tabControl1.SelectedTab = tabPage4;
                tCountShift.Text = "3";
                tCoefficientRepair.Text = "0,95";
                tCoefficientLight.Text = "0,95";
                tCountHourLight.Text = "6750";
                tPower.Text = "100";
                tProcentLosses.Text = "5";
                tCountLightPoint.Text = "364";
                tCoefficientLoad.Text = "1,5";
                tTariffLight.Text = "2,26";
                tAreaMachine.Text = "25";
                tCoefficientArea.Text = "1,3";

                button8.PerformClick();
                tabControl1.SelectedTab = tabPage5;

                tYearlyoutput.Text = "7500";
                tCostC1.Text = "180";
                tPercentCost2.Text = "1";

                tCoefficientEconomic.Text = "0,15";
                tInvestmenK1.Text = "500";
                tINvestmenK2.Text = "20500";







            }



        }
        private void Form5_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Diagnostics.Process.GetCurrentProcess().Kill();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            DataCa.addRow(dataGridView1, false);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataCa.addRemove(dataGridView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {

            int procent = int.Parse(textBox1.Text);
            double sum = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {

                double count = double.Parse(dataGridView1.Rows[i].Cells[2].Value.ToString());
                double price = double.Parse(dataGridView1.Rows[i].Cells[3].Value.ToString());
                sum += DataCa.MathRound(count * price);
                dataGridView1.Rows[i].Cells[4].Value = DataCa.MathRound(count * price);
            }

            DataCa.addRow(dataGridView1, false);
            DataCa.addRow(dataGridView1, false);
            DataCa.addRow(dataGridView1, false);
            int countDtg = dataGridView1.Rows.Count;



            dataGridView1.Rows[countDtg - 3].Cells[0].Value = "Разом";
            dataGridView1.Rows[countDtg - 3].Cells[4].Value = sum;
            dataGridView1.Rows[countDtg - 2].Cells[0].Value = "Транспортно-заготівельні витрати";
            dataGridView1.Rows[countDtg - 2].Cells[1].Value = "%";
            dataGridView1.Rows[countDtg - 2].Cells[2].Value = procent;
            dataGridView1.Rows[countDtg - 2].Cells[4].Value = sum * procent / 100;
            dataGridView1.Rows[countDtg - 1].Cells[0].Value = "Всього";
            dataGridView1.Rows[countDtg - 1].Cells[4].Value = DataCa.MathRound(sum + (sum * procent / 100));





        }

        private void button5_Click(object sender, EventArgs e)
        {
            DataCa.addRemove(dataGridView2);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DataCa.addRow(dataGridView2, false);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            double sum = 0;
            int sumJob = 0;
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {

                int count = int.Parse(dataGridView2.Rows[i].Cells[2].Value.ToString());
                double price = double.Parse(dataGridView2.Rows[i].Cells[1].Value.ToString());
                sum += DataCa.MathRound(count * price);
                sumJob += count;
                dataGridView2.Rows[i].Cells[3].Value = DataCa.MathRound(count * price);
            }

            dataGridView2.Rows.Add(2);
            int countDtg = dataGridView2.Rows.Count;
            dataGridView2.Rows[countDtg - 2].Cells[0].Value = "Всього";
            dataGridView2.Rows[countDtg - 2].Cells[2].Value = sumJob;
            dataGridView2.Rows[countDtg - 2].Cells[3].Value = sum;
            dataGridView2.Rows[countDtg - 1].Cells[0].Value = "Кошторисна (пряма) заробітна плата електриків-ремонтників за виконання робіт";
            dataGridView2.Rows[countDtg - 1].Cells[3].Value = sum;
        }

        private void button7_Click(object sender, EventArgs e)
        {

            PercentSalary = int.Parse(tPercentSalary.Text);
            PercentEsv = int.Parse(tPercentEsv.Text);
            PercentAdditional = int.Parse(tPercentAdditional.Text);
            PercentAdministrative = int.Parse(tPercentAdministrative.Text);
            PercentGeneral = int.Parse(tPercentGeneral.Text);
            PercentSum = int.Parse(tPercentSum.Text);
            int countDtg = dataGridView2.Rows.Count;
            double SalaryPr = double.Parse(dataGridView2.Rows[countDtg - 1].Cells[3].Value.ToString());

            PremiumSlary = DataCa.MathRound((SalaryPr * PercentSalary) / 100);

            MainSalary = DataCa.MathRound(SalaryPr + PremiumSlary);
            SalaryAdditional = DataCa.MathRound((MainSalary * PercentAdditional) / 100);
            Esv = DataCa.MathRound((PercentEsv * (MainSalary + SalaryAdditional)) / 100);

            GeneralSalary = DataCa.MathRound(PercentGeneral * (MainSalary) / 100);
            AdministrativeSalaty = DataCa.MathRound(PercentAdministrative * (MainSalary) / 100);

            double allDgt1 = double.Parse(dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[4].Value.ToString());

            FullCost = DataCa.MathRound(allDgt1 + MainSalary + SalaryAdditional + Esv + GeneralSalary + AdministrativeSalaty);


            SumProfit = DataCa.MathRound((PercentSum * FullCost) / 100);
            CostsWorks = DataCa.MathRound(SumProfit + FullCost);



            dataGridView3.Rows[0].Cells[2].Value = dataGridView1.Rows[dataGridView1.RowCount - 3].Cells[4].Value.ToString();

            dataGridView3.Rows[1].Cells[2].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[4].Value.ToString();
            dataGridView3.Rows[1].Cells[1].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[2].Value.ToString();

            dataGridView3.Rows[2].Cells[2].Value = dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[4].Value.ToString();
            dataGridView3.Rows[3].Cells[2].Value = MainSalary;

            dataGridView3.Rows[4].Cells[1].Value = PercentAdditional;
            dataGridView3.Rows[4].Cells[2].Value = SalaryAdditional;

            dataGridView3.Rows[5].Cells[1].Value = PercentEsv;
            dataGridView3.Rows[5].Cells[2].Value = Esv;

            dataGridView3.Rows[6].Cells[1].Value = PercentGeneral;
            dataGridView3.Rows[6].Cells[2].Value = GeneralSalary;

            dataGridView3.Rows[7].Cells[1].Value = PercentAdministrative;
            dataGridView3.Rows[7].Cells[2].Value = AdministrativeSalaty;


            dataGridView3.Rows[8].Cells[2].Value = FullCost;

            dataGridView3.Rows[9].Cells[1].Value = PercentSum;
            dataGridView3.Rows[9].Cells[2].Value = SumProfit;


            dataGridView3.Rows[10].Cells[2].Value = CostsWorks;

            profitAll = SumProfit;
            AllCosts = FullCost;

            DataCa.Ta[6] = CostsWorks;
            DataCa.Ta[15] = SumProfit;
            DataCa.Ta[16] = PercentSum;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            CountShift = double.Parse(tCountShift.Text);
            CoefficientRepair = double.Parse(tCoefficientRepair.Text);
            CoefficientLight = double.Parse(tCoefficientLight.Text);
            CountHourLight = double.Parse(tCountHourLight.Text);
            Power = double.Parse(tPower.Text);
            ProcentLosses = double.Parse(tProcentLosses.Text);

            CoefficientLoad = double.Parse(tCoefficientLoad.Text);
            TariffLight = double.Parse(tTariffLight.Text);
            AreaMachine = double.Parse(tAreaMachine.Text);
            CoefficientAreaMain = double.Parse(tCoefficientArea.Text);
            double CountMachine = double.Parse(tCountMachine.Text);
            dataGridView4.Rows.Clear();
            int dtgf1 = DataCa.f2.MyDataGridView.RowCount;

            for (int i = 0; i < dtgf1; i++)
            {
                string name = DataCa.f2.MyDataGridView.Rows[i].Cells[1].Value.ToString();
                int count = int.Parse(DataCa.f2.MyDataGridView.Rows[i].Cells[2].Value.ToString());
                double P = double.Parse(DataCa.f2.MyDataGridView.Rows[i].Cells[3].Value.ToString());
                double Kv = double.Parse(DataCa.f2.MyDataGridView.Rows[i].Cells[4].Value.ToString());
                dataGridView4.Rows.Add(name, count, P, Kv);

            }

            int AllPerson = 0, SumW = 0;
            for (int i = 0; i < dtgf1; i++)
            {
                int count = int.Parse(dataGridView4.Rows[i].Cells[1].Value.ToString());
                AllPerson += count;
            }
            AreaMain = (int)DataCa.MathRoundInt(AreaMachine * CountMachine);
            AreaGeneral = (int)DataCa.MathRoundInt(AreaMain * CoefficientAreaMain);
            CountLightPoint = double.Parse(tCountLightPoint.Text) * AreaGeneral / Power;
            double PnomLight = DataCa.MathRoundOne(((Power * CountLightPoint * CoefficientLight * CoefficientLoad) / (1000 * (1 - (ProcentLosses / 100)))));

            dataGridView4.Rows.Add("Освітлення", "0", PnomLight, CoefficientLight);
            dataGridView4.Rows.Add("Всього");
            AllPerson = 0; SumW = 0;
            for (int i = 0; i < dtgf1; i++)
            {
                double MainFond = DataCa.MathRoundInt(((DataCa.daysCalendar - DataCa.dayHoliday - DataCa.dayWekkend) * DataCa.shiftDuration - DataCa.dayPreHoliday) * CountShift * CoefficientRepair);
                dataGridView4.Rows[i].Cells[4].Value = MainFond;
                int count = int.Parse(dataGridView4.Rows[i].Cells[1].Value.ToString());
                double P = double.Parse(dataGridView4.Rows[i].Cells[2].Value.ToString());
                double Kv = double.Parse(dataGridView4.Rows[i].Cells[3].Value.ToString());
                int W = int.Parse((DataCa.MathRoundInt((count * P * Kv * MainFond))).ToString());
                dataGridView4.Rows[i].Cells[5].Value = W;
                SumW += W;

                AllPerson += count;
            }



            dataGridView4.Rows[dataGridView4.RowCount - 2].Cells[4].Value = CountHourLight;
            double LightW = DataCa.MathRoundInt(
                double.Parse(dataGridView4.Rows[dataGridView4.RowCount - 2].Cells[2].Value.ToString()) *
                 double.Parse(dataGridView4.Rows[dataGridView4.RowCount - 2].Cells[4].Value.ToString()) *
                double.Parse(dataGridView4.Rows[dataGridView4.RowCount - 2].Cells[3].Value.ToString()));
            dataGridView4.Rows[dataGridView4.RowCount - 2].Cells[5].Value = LightW;

            dataGridView4.Rows[dataGridView4.RowCount - 1].Cells[1].Value = AllPerson;
            dataGridView4.Rows[dataGridView4.RowCount - 1].Cells[5].Value = DataCa.MathRoundInt((SumW + LightW));

            CostElectricity = DataCa.MathRound(LightW * TariffLight);



            MainSalaryByElectricity = DataCa.MathRound(DataCa.MathRoundInt((SumW + LightW)) * TariffLight);
            label2.Text = "Вартість  річної електроенергії для освітлення  Вріч.е/е.о грн. - " + CostElectricity +
            "\nПитомий показник площі на один верстат Sпит.д - " + AreaMain +
            "\nЗагальна площа Sзаг - " + AreaGeneral +
            "\nОсновна оплата за заявлену (спожиту) активну електроенергію, грн. - " + MainSalaryByElectricity;


            DataCa.Ta[13] = (SumW + LightW);
            DataCa.Ta[14] = TariffLight;
            DataCa.Ta[17] = AreaGeneral;
        }



        private void InputNumber(object sender, KeyPressEventArgs e)
        {
            DataCa.TextFloat_KeyPress(sender, e);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            dataGridView5.Rows.Clear();
            string[,] data3 = new string[,]
              {
                    {"Річна трудомісткість ремонтних робіт,  н-год, в тому числі:", ""},
                    {"капітальних ремонтів", "н-год"},
                    {"поточних ремонтів", "н-год"},
                    {"Витрати на ремонт, грн. в тому числі:", ""},
                    {"капітальний ", "грн"},
                    {"поточний", "грн"},
                    {"Вартість робіт за внутрішніми планово-розрахунковими цінами, ", "грн"},
                    {"Чисельність персоналу електриків-ремонтників", "чол"},
                    {"Чисельність персоналу чергових електриків", "чол"},
                    {"Фонд оплати праці електриків-ремонтників", "грн"},
                    {"Фонд оплати праці чергових електриків", "грн"},
                    {"Середньомісячна заробітна плата робітника, електриків-ремонтників", "грн./особу"},
                    {"Середньомісячна заробітна плата робітника, чергових електриків", "грн./особу"},
                    {"Кількість витраченої електроенергії", "кВт."},
                    {"Вартість одного кВт витраченої електроенергії", "грн"},
                    {"Нормативний прибуток від виконання робіт", "грн"},
                    {"Рентабельність  робіт", "%"},
                    {"Виробнича площа дільниці", "М2"},
                    {"Продуктивність праці", "грн/особу"},
                    {"Продуктивність праці", "н-год/особу"},
                    {"Річний економічний ефект від впровадження  заходів", "грн"},
                    {"Термін окупності додаткових капіталовкладень", "роки"},
              };



            // Добавьте строки из массива в DataGridView
            for (int i = 0; i < data3.GetLength(0); i++)
            {
                dataGridView5.Rows.Add(data3[i, 0], data3[i, 1]);
            }

            double YearlyEconomic, YearlyEffect, PaybackPeriod;
            double YearlyOutput, CostC1, CostC2, CoefficientEconomic, InvestmenK1, InvestmenK2;


            YearlyOutput = double.Parse(tYearlyoutput.Text);
            CostC1 = double.Parse(tCostC1.Text);
            CostC2 = DataCa.MathRound(CostC1 - double.Parse(tPercentCost2.Text) * MathRound(CostC1 / 100));

            CoefficientEconomic = double.Parse(tCoefficientEconomic.Text);
            InvestmenK1 = double.Parse(tInvestmenK1.Text);
            InvestmenK2 = double.Parse(tINvestmenK2.Text);


            YearlyEconomic = DataCa.MathRound(CostC1 - CostC2) * YearlyOutput;
            YearlyEffect = DataCa.MathRound(
                (DataCa.MathRound(CostC1 - CostC2) -
                DataCa.MathRound(CoefficientEconomic *
                DataCa.MathRound((InvestmenK2 - InvestmenK1) / YearlyOutput))) * YearlyOutput
                );

            PaybackPeriod = DataCa.MathRound((InvestmenK2 - InvestmenK1) / YearlyEffect);

            int Re = (int)DataCa.MathRoundInt(profitAll / AllCosts * 100);

            label3.Text = "Умовно річний економічний ефект Еум.річ грн. - " + YearlyEconomic +
                "\nРічний економічний ефект Еріч грн. " + YearlyEffect +
                 "\nТермін окупності додаткових капіталовкладень на \nвпровадження заходів по зниженню собівартості Ток  " + PaybackPeriod + "р. ≈ " + DataCa.MathRound(PaybackPeriod * 12) + "міс." +
                 "\nРентабельність робіт R - " + Re + "%";
            DataCa.Ta[20] = YearlyEffect;
            DataCa.Ta[21] = PaybackPeriod;

            for (int i = 0; i < dataGridView5.RowCount; i++)
            {
                dataGridView5.Rows[i].Cells[2].Value = DataCa.Ta[i];
            }



        }

        private void button10_Click(object sender, EventArgs e)
        {

            string text = "Таблиця 5.5.2 – Кошторис  вартості робіт з монтажу електрообладнання";
            int[] columnWidths = { 200, 30, 30, 70 };
            DataCa.PrintDtg(dataGridView2, text, columnWidths, 40, true, false);

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            string text = "Таблиця 5.5.1 -  Кошторис витрат на придбання електрообладнання";
            int[] columnWidths = { 200, 30, 30, 50, 70 };
            DataCa.PrintDtg(dataGridView1, text, columnWidths, 40, true, false);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            string text = "Таблиця 5.5.3 - Кошторис витрат на придбання і монтаж електрообладнання";
            int[] columnWidths = { 200, 30, 50 };
            DataCa.PrintDtg(dataGridView3, text, columnWidths, 40, true, false);
        }

        private void button13_Click(object sender, EventArgs e)
        {

            string text = "Таблиця 5.6.1 - Споживання електроенергії на дільниці";
            int[] columnWidths = { 200, 30, 30, 30, 50, 50 };
            DataCa.PrintDtg(dataGridView4, text, columnWidths, 40, true, false);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            string text = "Таблиця 8.1 – Техніко - економічні показники ";
            int[] columnWidths = { 200, 30, 60 };
            DataCa.PrintDtg(dataGridView5, text, columnWidths, 40, true, false);
        }
    }
}
