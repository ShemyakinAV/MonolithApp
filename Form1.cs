using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Monolith
{
    public partial class General_Form : Form
    {
        public int[] start_serverPrices = new int[] { 1000, 800, 660, 1640, 1000, 900 };
       

        public General_Form()
        {
            InitializeComponent();
            A_buttonNext.Enabled = false;
        }

        private void A_buttonNext_Click(object sender, EventArgs e)
        {
            //Создание .xlsx-файла
            string filePath = "D:\\test.xlsx";
            if (File.Exists(filePath))
                File.Delete(filePath);
            Excel.Application oApp;
            Excel.Workbook oBook;
            Excel.Worksheet oSheet;
            Excel.Range oRange;
            oApp = new Excel.Application();
            oBook = oApp.Workbooks.Add();
            oSheet = (Excel.Worksheet)oBook.Worksheets.get_Item(1);    
            oSheet.Cells[1, 1] = "Название";
            oSheet.Cells[1, 2] = "Количество";
            oSheet.Cells[1, 3] = "Сумма";
            //
            int LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; //Поиск последней заполненной строки
           
            if(A1_count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] =A1_IPtelServer.Text;
                oSheet.Cells[LastRow+1, 2] = A1_count.Value;
                oSheet.Cells[LastRow+1, 3] = A1_count.Value*start_serverPrices[0];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (A2_count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = A2_messageServer.Text;
                oSheet.Cells[LastRow + 1, 2] = A2_count.Value;
                oSheet.Cells[LastRow + 1, 3] = A2_count.Value * start_serverPrices[1];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (A3_count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = A3_PostServer.Text;
                oSheet.Cells[LastRow + 1, 2] = A3_count.Value;
                oSheet.Cells[LastRow + 1, 3] = A3_count.Value * start_serverPrices[2];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (A4_count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = A4_remoteServer.Text;
                oSheet.Cells[LastRow + 1, 2] = A4_count.Value;
                oSheet.Cells[LastRow + 1, 3] = A4_count.Value * start_serverPrices[3];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (A5_count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = A5_proxyServer.Text;
                oSheet.Cells[LastRow + 1, 2] = A5_count.Value;
                oSheet.Cells[LastRow + 1, 3] = A5_count.Value * start_serverPrices[4];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (A6_count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = A6_reserveServer.Text;
                oSheet.Cells[LastRow + 1, 2] = A6_count.Value;
                oSheet.Cells[LastRow + 1, 3] = A6_count.Value * start_serverPrices[5];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            for (int i = LastRow; i >= 1; i--)
            {

                oSheet.Rows[i].HorizontalAlignment = -4108;
                oSheet.Rows[i].VerticalAlignment = -4108;
                if ((i % 2) == 0)
                {
                    var columnHeadingsRange = oSheet.Range[
                        oSheet.Cells[i, 1],
                        oSheet.Cells[i, 3]];

                    columnHeadingsRange.Interior.Color = Color.FromArgb(0, 241, 242, 242);



                }
                else
                {
                    var columnHeadingsRange = oSheet.Range[
                        oSheet.Cells[i, 1],
                        oSheet.Cells[i, 3]];

                    columnHeadingsRange.Interior.Color = Color.FromArgb(0, 248, 249, 248);
                }
                if (oSheet.Cells[i, 2].Text == "0" || oSheet.Cells[i, 2].Value == null)
                {
                    oSheet.Rows[i].Delete();
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                }
            }
            var FirstInteriorRange = oSheet.Range[oSheet.Cells[1, 1],
                oSheet.Cells[1, 3]];
            FirstInteriorRange.Interior.Color = Color.FromArgb(0, 214, 209, 235);
            var LastInteriorRange = oSheet.Range[
                oSheet.Cells[LastRow + 1, 1],
                oSheet.Cells[LastRow + 1, 3]];
            LastInteriorRange.Interior.Color = Color.FromArgb(0, 214, 209, 235);
            Console.WriteLine(LastRow);
            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();
            oBook.SaveAs(filePath);
            oBook.Close();
            oApp.Quit();
            Application.Exit();
            A_panel.Hide();
            A_panelExtra.Hide();
            hardwarePanel.Show();
        }
        
         private void A1_count_ValueChanged(object sender, EventArgs e)
        {
            if (A1_count.Value != 0|| A2_count.Value != 0|| A3_count.Value != 0|| A4_count.Value != 0|| A5_count.Value != 0|| A6_count.Value != 0)
                A_buttonNext.Enabled = true;
            else
                A_buttonNext.Enabled = false;
        }

        private void A2_count_ValueChanged(object sender, EventArgs e)
        {
            if (A2_count.Value != 0 || A1_count.Value != 0 || A3_count.Value != 0 || A4_count.Value != 0 || A5_count.Value != 0 || A6_count.Value != 0)
                A_buttonNext.Enabled = true;
            else
                A_buttonNext.Enabled = false;
        }

        private void A3_count_ValueChanged(object sender, EventArgs e)
        {
            if (A2_count.Value != 0 || A1_count.Value != 0 || A3_count.Value != 0 || A4_count.Value != 0 || A5_count.Value != 0 || A6_count.Value != 0)
                A_buttonNext.Enabled = true;
            else
                A_buttonNext.Enabled = false;
        }

        private void A4_count_ValueChanged(object sender, EventArgs e)
        {
            if (A2_count.Value != 0 || A1_count.Value != 0 || A3_count.Value != 0 || A4_count.Value != 0 || A5_count.Value != 0 || A6_count.Value != 0)
                A_buttonNext.Enabled = true;
            else
                A_buttonNext.Enabled = false;
        }

        private void A5_count_ValueChanged(object sender, EventArgs e)
        {
            if (A2_count.Value != 0 || A1_count.Value != 0 || A3_count.Value != 0 || A4_count.Value != 0 || A5_count.Value != 0 || A6_count.Value != 0)
                A_buttonNext.Enabled = true;
            else
                A_buttonNext.Enabled = false;
        }

        private void A6_count_ValueChanged(object sender, EventArgs e)
        {
            if (A2_count.Value != 0 || A1_count.Value != 0 || A3_count.Value != 0 || A4_count.Value != 0 || A5_count.Value != 0 || A6_count.Value != 0)
                A_buttonNext.Enabled = true;
            else
                A_buttonNext.Enabled = false;
        }
