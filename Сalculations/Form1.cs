namespace Сalculations
{
    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataCa.f3.Show();
            this.Hide();
           /* double YearlyEconomic, YearlyEffect, PaybackPeriod;
            double YearlyOutput, CostC1, CostC2, CoefficientEconomic, InvestmenK1, InvestmenK2;
            YearlyOutput = double.Parse("56");
            CostC1 = double.Parse("39930,67");
            CostC2 = DataCa.MathRound(CostC1 - double.Parse("1") * DataCa.MathRound(CostC1 / 100));

            CoefficientEconomic = double.Parse("0,15");
            InvestmenK1 = double.Parse("0");
            InvestmenK2 = double.Parse("20000");


            YearlyEconomic = DataCa.MathRound(CostC1 - CostC2) * YearlyOutput;
            YearlyEffect = DataCa.MathRound(
                (DataCa.MathRound(CostC1 - CostC2) -
                DataCa.MathRound(CoefficientEconomic *
                DataCa.MathRound((InvestmenK2 - InvestmenK1) / YearlyOutput))) * YearlyOutput
                );
            MessageBox.Show(DataCa.MathRound(CoefficientEconomic * DataCa.MathRound((InvestmenK2 - InvestmenK1) / YearlyOutput)).ToString());
            MessageBox.Show(DataCa.MathRound(CostC1 - CostC2).ToString());
            MessageBox.Show((DataCa.MathRound(CostC1 - CostC2) - DataCa.MathRound(CoefficientEconomic * DataCa.MathRound((InvestmenK2 - InvestmenK1) / YearlyOutput))).ToString());
            MessageBox.Show(YearlyEffect.ToString());*/
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Diagnostics.Process.GetCurrentProcess().Kill();

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}