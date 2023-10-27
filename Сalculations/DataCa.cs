using DGVPrinterHelper;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Сalculations
{
    static internal class DataCa
    {

        public static string[] name_1;
        public static int[] count_1;
        public static double[] Ph;
        public static double[] Kv;
        public static String[] categories_1;

        public static int[] Ttable_rc;

        public static int[] Ttabl1;
        public static double[] Bk1;
        public static double[] Br1;
          public static double[] Bb1;
        public static double[] Bo1;
        public static double[] Bc1;

        public static int[] Ttabl2;
        public static double[] Bk2;
        public static double[] Br2;
          public static double[] Bb2;
        public static double[] Bo2;
        public static double[] Bc2;
        public static int[] period;

        public static double[] hard = { 16, 26, 16, 14, 18, 13, 8, 12, 10, 19, 14, 7 };
        public static int Trich;
        public static double sumHard;
        public static double rAvarage,tPot,tKap;
        public static int daysCalendar,dayHoliday, dayPreHoliday, dayWekkend,shiftDuration;

        public static void TextFloat_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != ','))
            {
                e.Handled = true;
            }

           
            if ((e.KeyChar == ',') && ((sender as TextBox).Text.IndexOf(',') > -1))
            {
                e.Handled = true;
            }
        }
        public static double[] Ta = new double[22];

            
        


   
    public static Form2 f2 = new Form2();
        public static Form3 f3 = new Form3();
        public static Form4 f4 = new Form4();
        public static Form5 f5 = new Form5();
   


   
        public static CommonMenu commonMenu;
        public static void addRow(DataGridView dtg, Boolean number = false)
        {
            dtg.Rows.Add(1);

            int count = dtg.Rows.Count;
            if (number) { dtg.Rows[count - 1].Cells[0].Value = dtg.Rows.Count; }

        }
        public static void addRemove(DataGridView dtg)
        {
            int count = dtg.Rows.Count;
            if (count != 0) { dtg.Rows.RemoveAt(count - 1); }

        }
        public static double MathRound(double number)
        { 
            return Math.Round(number * 100, MidpointRounding.AwayFromZero) / 100;
        }


        public static double MathRoundOne(double number)
        {
            return Math.Round(number * 10, MidpointRounding.AwayFromZero) / 10;
        }
        
        public static double MathRoundInt(double number)
        {
            return Math.Round(number, MidpointRounding.AwayFromZero);
        }
        public static void PrintDtg(DataGridView dtg, string title, int[] arrayWidth, int HeightHeader, Boolean IsNeed, Boolean Land)
        {
            DGVPrinter printer = new DGVPrinter();
            printer.Title = title;
            printer.PageNumbers = true;
            printer.PageNumberInHeader = false;
            printer.PorportionalColumns = true;
            printer.PageSettings.Landscape = Land;


            printer.HeaderCellAlignment = StringAlignment.Center;
            printer.Footer = DateTime.Now.ToString("dd.MM.yyyy");
            printer.FooterSpacing = 15;

            if (IsNeed)
            {
                int[] oldColumnWidths = new int[dtg.ColumnCount];
                int[] oldColumnMinWidths = new int[dtg.ColumnCount];
                string[] oldColumnSize = new string[dtg.ColumnCount];

                int oldHeader = dtg.ColumnHeadersHeight;
                for (int i = 0; i < dtg.ColumnCount; i++)
                {
                    oldColumnMinWidths[i] = dtg.Columns[i].MinimumWidth;
                    oldColumnWidths[i] = dtg.Columns[i].Width;
                    oldColumnSize[i] = dtg.Columns[i].AutoSizeMode.ToString();
                    dtg.Columns[i].MinimumWidth = 10;
                    dtg.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet;
                    dtg.Columns[i].Width = arrayWidth[i];
                    dtg.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dtg.ColumnHeadersHeight = HeightHeader;
                }

                object[,] originalData = new object[dtg.Rows.Count, dtg.Columns.Count];

                for (int i = 0; i < dtg.Rows.Count; i++)
                {
                    for (int j = 0; j < dtg.Columns.Count; j++)
                    {
                        originalData[i, j] = dtg.Rows[i].Cells[j].Value;
                    }
                }
                foreach (DataGridViewRow row in dtg.Rows)
                {
                    dtg.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        cell.Style.WrapMode = DataGridViewTriState.True;

                    }

                }




                printer.RowHeight = DGVPrinter.RowHeightSetting.CellHeight;
                printer.PrintDataGridView(dtg);
                dtg.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

                dtg.AutoResizeRows();
                for (int i = 0; i < dtg.Rows.Count; i++)
                {
                    for (int j = 0; j < dtg.Columns.Count; j++)
                    {
                        dtg.Rows[i].Cells[j].Value = originalData[i, j];
                    }
                }
                for (int i = 0; i < dtg.ColumnCount; i++)
                {
                    dtg.Columns[i].MinimumWidth = oldColumnMinWidths[i];
                    dtg.Columns[i].Width = oldColumnWidths[i];
                    dtg.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dtg.ColumnHeadersHeight = oldHeader;
                    if (Enum.TryParse(oldColumnSize[i], out DataGridViewAutoSizeColumnMode mode))
                    {
                        dtg.Columns[i].AutoSizeMode = mode;
                    }
                   
                  
                  
                }

            }
            else
            {

                printer.PrintDataGridView(dtg);
            }

        }

        public class CommonMenu
        {
            public ToolStripMenuItem MenuMenu { get; private set; }
           
            private Form currentForm;
            public CommonMenu(Form form)
            {
                currentForm = form;
                MenuMenu = new ToolStripMenuItem("Меню");
               
                MenuMenu.DropDownItems.Add(new ToolStripMenuItem("Меню", null, MenuMenu_Click));
                MenuMenu.DropDownItems.Add(new ToolStripMenuItem("Організація виробництва", null, MenuFirst_Click));
                MenuMenu.DropDownItems.Add(new ToolStripMenuItem("Економічна частина 1",null, Ec1));
                MenuMenu.DropDownItems.Add(new ToolStripMenuItem("Економічна частина 2",null, Ec2));

    

         
            }

          
            private void MenuMenu_Click(object sender, EventArgs e)
            {
                currentForm.Hide();
                f3.Show();
            }
            private void MenuFirst_Click(object sender, EventArgs e)
            {
                currentForm.Hide();
                f2.Show();
            }
            private void Ec1(object sender, EventArgs e)
            {
                currentForm.Hide();
                f4.Show();
            }
            private void Ec2(object sender, EventArgs e)
            {
                currentForm.Hide();
                f5.Show();
            }

        }










    }

}
