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
    public partial class GeneralForm : Form
    {
        public int[] start_serverPrices = new int[] { 1000, 800, 660, 1640, 1000, 900 };
        public int[] start_extraPrices = new int[] { 500, 1000, 2000, 1200, 3000, 1500 };
        public int[] start_phonePrices = new int[] { 620, 1220, 1580, 50 };
        public int[] start_mothPayPrices = new int[] { 150, 200, 13, 10,13 };
        public int[] start_pcPrices = new int[] { 1000, 35, 100, 360,120 };
        public int start_serverRenPrices = 65;
        public int start_serverServicePrice = 60;
        public int currentserverRenPrices;

        public string filePath = "C:\\MonolithPlus\\test.xlsx";
        
       
        public GeneralForm()
        {
            InitializeComponent();
            hardwarePanel.Hide();
            phonePanel.Hide();
            A_panelExtra.Hide();
            monthpay_phonePanel.Hide();
            pcChoosingPanel.Hide();
            rentServer_Panel.Hide();
          
            Aextra_doneButton.Enabled = false;
            phoneDoneButton.Enabled = false;
            
        }

        private void A_buttonNext_Click(object sender, EventArgs e)
        {
            this.Text = "Аппаратное обеспечение";
            A_panel.Hide();
            A_panelExtra.Hide();
            hardwarePanel.Show();
        }
        
        private void A_buttonExtra_Click(object sender, EventArgs e)
        {
            this.Text = "Серверная структура (дополнительные сервисы)";
            A_panel.Hide();
            A_panelExtra.Show();
        }

        private void Aextra_backButton_Click(object sender, EventArgs e)
        {
            this.Text = "Серверная структура";
            A1extra_count.Value = A2extra_count.Value = A3extra_count.Value=A4extra_count.Value=A5extra_count.Value =0;
            A_panel.Show();
            A_panelExtra.Hide();          
        }

        private void Aextra_doneButton_Click(object sender, EventArgs e)
        {
            this.Text = "Серверная структура";
            A_panel.Show();
            A_panelExtra.Hide();
        }

        private void HardwareDoneButton_Click(object sender, EventArgs e)
        {
            rentServerCount2.Value = A1_count.Value + A2_count.Value + A3_count.Value + A4_count.Value + A5_count.Value + A6_count.Value;
            this.Text = "Ежемесячные расходы на сервера и услуги";
            rentServerCount1.Value = A1_count.Value + A2_count.Value + A3_count.Value + A4_count.Value + A5_count.Value + A6_count.Value;
            hardwarePanel.Hide();
            rentServer_Panel.Show();
            rentServerprice_Text.Text = rentServerCount1.Value * start_serverRenPrices + "$";
            
        }

        private void HardwareBackButton_Click(object sender, EventArgs e)
        {
            this.Text = "Серверная структура";
            hardwarePanel.Hide();
            A_panel.Show();
        }

        private void HardwarePCButton_Click(object sender, EventArgs e)
        {
            this.Text = "Аппаратное обеспечение - компьютеры";
            hardwarePanel.Hide();
            pcChoosingPanel.Show();
        }

        private void HardwarePhonesButton_Click(object sender, EventArgs e)
        {
            this.Text = "Аппаратное обеспечение - телефоны";
            hardwarePanel.Hide();
            phonePanel.Show();
        }

        private void PhoneDoneButton_Click(object sender, EventArgs e)
        {
            this.Text = "Ежемесячные расходы на телефонию";
            monthpay_phonePanel.Show();
            phonePanel.Hide();
        }

        private void MonthPhone_doneButton_Click(object sender, EventArgs e)
        {
            this.Text = "Аппаратное обеспечение";
            monthpay_phonePanel.Hide();
            hardwarePanel.Show();
            phoneChoose_Text.Text = null;
             if (phone1Count.Value > 0)
            {
                phoneChoose_Text.AppendText(Phone1_Text.Text + " Количество: " + phone1Count.Value.ToString() + " Стоимость: " + phone1Count.Value * start_phonePrices[0] + "\n");
            }
            if (phone2Count.Value > 0)
            {
                phoneChoose_Text.AppendText(Phone2_Text.Text + " Количество: " + phone2Count.Value.ToString() + " Стоимость: " + phone2Count.Value * start_phonePrices[1] + "\n");
            }
            if (phone3Count.Value > 0)
            {
                phoneChoose_Text.AppendText(Phone3_Text.Text + " Количество: " + phone3Count.Value.ToString() + " Стоимость: " + phone3Count.Value * start_phonePrices[2] + "\n");
            }
            if (phone4Count.Value > 0)
            {
                phoneChoose_Text.AppendText(Phone4_Text.Text + " Количество: " + phone4Count.Value.ToString() + " Стоимость: " + phone4Count.Value * start_phonePrices[3] + "\n");
            }
            if (month1Price.Value > 0)
            {
                phoneChoose_Text.AppendText(monthpay1_text.Text + " Количество: " + month1Price.Value.ToString() + " Стоимость: " + month1Price.Value * start_mothPayPrices[0] + "\n");
            }
            if (month2Price.Value > 0)
            {
                phoneChoose_Text.AppendText(monthpay2_text.Text + " Количество: " + month2Price.Value.ToString() + " Стоимость: " + month2Price.Value * start_mothPayPrices[1] + "\n");
            }
            if (month3Price.Value > 0)
            {
                phoneChoose_Text.AppendText(monthpay3_text.Text + " Количество: " + month3Price.Value.ToString() + " Стоимость: " + month3Price.Value * start_mothPayPrices[2] + "\n");
            }
            if (month4Price.Value > 0)
            {
                phoneChoose_Text.AppendText(monthpay4_text.Text + " Количество: " + month4Price.Value.ToString() + " Стоимость: " + month4Price.Value * start_mothPayPrices[3] + "\n");
            }
            if (isCompanyPay_ChBox.Checked == false)
            {
                month1Price.Value = 0;
                month2Price.Value = month2Price.Minimum;

            }
        }
      
    

        private void PcChoosingDone_Button_Click(object sender, EventArgs e)
        {
            this.Text = "Аппаратное обеспечение";
            pcChoosingPanel.Hide();
            hardwarePanel.Show();
            pcChoose_Text.Text = null;
            if (pcChoosingPanelCount1.Value > 0)
            {
                pcChoose_Text.AppendText(pcChoose1_Text.Text + ": " + pcChoosingPanelCount1.Value.ToString() + "\n");
            }
            if (pcChoosingPanelCount2.Value > 0)
            {
                pcChoose_Text.AppendText(pcChoose2_Text.Text + " Количество: " + pcChoosingPanelCount2.Value.ToString() + " Стоимость: " + pcChoosingPanelCount2.Value * start_pcPrices[0] + "\n");
            }
            if (pcChoosingPanelCount3.Value > 0 || pcChoosingPanelCount1.Value > 0)
            {
                pcChoose_Text.AppendText(pcChoose3_Text.Text + " Количество: " + (pcChoosingPanelCount1.Value).ToString() + " Стоимость: " + pcChoosingPanelCount1.Value * start_pcPrices[1] + "\n");
            }
            if (pcChoosingPanelCount4.Value > 0|| pcChoosingPanelCount1.Value > 0)
            {
                pcChoose_Text.AppendText(pcChoose4_Text.Text + " Количество: " + (pcChoosingPanelCount4.Value+ pcChoosingPanelCount1.Value).ToString() + " Стоимость: " + (pcChoosingPanelCount4.Value+ pcChoosingPanelCount1.Value) * start_pcPrices[2] + "\n");
            }
            if (pcChoosingPanelCount5.Value > 0)
            {
                pcChoose_Text.AppendText(isTripleEncryptionNeed_ChBox.Text + " Количество: " + pcChoosingPanelCount5.Value.ToString() + " Стоимость: " + pcChoosingPanelCount5.Value * start_pcPrices[3] + "\n");
            }
            if (pcChoosingPanelCount6.Value > 0)
            {
                pcChoose_Text.AppendText(isProgrammBlockNeed_ChBox.Text + " Количество: " + pcChoosingPanelCount6.Value.ToString() + " Стоимость: " + pcChoosingPanelCount6.Value * start_pcPrices[4] + "\n");
            }
            pcChoosingPanelCount3.Value = pcChoosingPanelCount4.Value = pcChoosingPanelCount1.Value;
        }

        private void PcChoosingBack_Button_Click(object sender, EventArgs e)
        {
            this.Text = "Аппаратное обеспечение";
            pcChoosingPanel.Hide();
            hardwarePanel.Show();
        }

        private void PhoneBackButton_Click(object sender, EventArgs e)
        {
            this.Text = "Аппаратное обеспечение";
            hardwarePanel.Show();
            phonePanel.Hide();
        }


        private void MonthPhoneBack_Button_Click(object sender, EventArgs e)
        {
            this.Text = "Аппаратное обеспечение - телефоны";
            monthpay_phonePanel.Hide();
            phonePanel.Show();
        }

        private void RentServerBack_Button_Click(object sender, EventArgs e)
        {
            this.Text = "Аппаратное обеспечение";
            rentServer_Panel.Hide();
            hardwarePanel.Show();
        }

        private void Finish_Button_Click(object sender, EventArgs e)
        {

            if (isServiceNeed.Checked == false)
            {
                rentServerCount2.Value = 0;
            }

            //Создание .xlsx-файла

            if (File.Exists(filePath))
                File.Delete(filePath);
            Excel.Application oApp;
            Excel.Workbook oBook;
            Excel.Worksheet oSheet;
            oApp = new Excel.Application();
            oBook = oApp.Workbooks.Add();
            oSheet = (Excel.Worksheet)oBook.Worksheets.get_Item(1);

            //
            int LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; //Поиск последней заполненной строки
            if(A1_count.Value> 0 || A2_count.Value > 0|| A3_count.Value > 0 || A4_count.Value > 0 || A5_count.Value > 0 || A6_count.Value > 0)
            {
                oSheet.Cells[LastRow, 1] = "Серверная структура";
                oSheet.Cells[LastRow, 1].Font.Bold = true;
                oSheet.Cells[LastRow+1, 1] = "Название";
                oSheet.Cells[LastRow + 1, 2] = "Количество";
                oSheet.Cells[LastRow + 1, 3] = "Сумма";
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (A1_count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = A1_IPtelServer.Text;
                oSheet.Cells[LastRow + 1, 2] = A1_count.Value;
                oSheet.Cells[LastRow + 1, 3] = A1_count.Value * start_serverPrices[0];
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
            if (A1_count.Value > 0 || A2_count.Value > 0 || A3_count.Value > 0 || A4_count.Value > 0 || A5_count.Value > 0 || A6_count.Value > 0)
            {
                oSheet.Cells[LastRow+1, 2] = "Итого:";
                oSheet.Cells[LastRow+1, 3] = A1_count.Value * start_serverPrices[0]+ A2_count.Value * start_serverPrices[1]+ A3_count.Value * start_serverPrices[2]+ A4_count.Value * start_serverPrices[3]+ A5_count.Value * start_serverPrices[4]+ A6_count.Value * start_serverPrices[5];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
                if (A1extra_count.Value > 0|| A2extra_count.Value > 0 || A3extra_count.Value > 0 || A4extra_count.Value > 0 || A5extra_count.Value > 0 || A6extra_count.Value > 0)
            {

                if (LastRow - 1 != 0)
                {
                    oSheet.Cells[LastRow + 2, 1] = "Серверная структура (дополнительные сервисы)";
                    oSheet.Cells[LastRow + 2, 1].Font.Bold = true;
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    oSheet.Cells[LastRow + 1, 1] = "Название";
                    oSheet.Cells[LastRow + 1, 2] = "Количество";
                    oSheet.Cells[LastRow + 1, 3] = "Сумма";
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                }
                else
                {
                    oSheet.Cells[LastRow, 1] = "Серверная структура (дополнительные сервисы)";
                    oSheet.Cells[LastRow, 1].Font.Bold = true;
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    oSheet.Cells[LastRow + 1, 1] = "Название";
                    oSheet.Cells[LastRow + 1, 2] = "Количество";
                    oSheet.Cells[LastRow + 1, 3] = "Сумма";
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                }

            }
            if (A1extra_count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = Aextra_proxyServer.Text;
                oSheet.Cells[LastRow + 1, 2] = A1extra_count.Value;
                oSheet.Cells[LastRow + 1, 3] = A1extra_count.Value * start_extraPrices[0];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (A2extra_count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = Aextra_cloudStore.Text;
                oSheet.Cells[LastRow + 1, 2] = A2extra_count.Value;
                oSheet.Cells[LastRow + 1, 3] = A2extra_count.Value * start_extraPrices[1];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (A3extra_count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = Aextra_remoteControl.Text;
                oSheet.Cells[LastRow + 1, 2] = A3extra_count.Value;
                oSheet.Cells[LastRow + 1, 3] = A3extra_count.Value * start_extraPrices[2];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (A4extra_count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = Aextra_systemMonitor.Text;
                oSheet.Cells[LastRow + 1, 2] = A4extra_count.Value;
                oSheet.Cells[LastRow + 1, 3] = A4extra_count.Value * start_extraPrices[3];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (A5extra_count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = Aextra_techSupportSystem.Text;
                oSheet.Cells[LastRow + 1, 2] = A5extra_count.Value;
                oSheet.Cells[LastRow + 1, 3] = A5extra_count.Value * start_extraPrices[4];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (A6extra_count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = Aextra_rechargeSites.Text;
                oSheet.Cells[LastRow + 1, 2] = A6extra_count.Value;
                oSheet.Cells[LastRow + 1, 3] = A6extra_count.Value * start_extraPrices[5];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (A1extra_count.Value > 0 || A2extra_count.Value > 0 || A3extra_count.Value > 0 || A4extra_count.Value > 0 || A5extra_count.Value > 0 || A6extra_count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 2] = "Итого:";
                oSheet.Cells[LastRow + 1, 3] = A1extra_count.Value * start_extraPrices[0] + A2extra_count.Value * start_extraPrices[1] + A3extra_count.Value * start_extraPrices[2] + A4extra_count.Value * start_extraPrices[3] + A5extra_count.Value * start_extraPrices[4] + A6extra_count.Value * start_extraPrices[5];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (phone1Count.Value > 0 || phone2Count.Value > 0 || phone3Count.Value > 0 || phone4Count.Value > 0)
            {

                if (LastRow - 1 != 0)
                {
                    oSheet.Cells[LastRow + 2, 1] = "Аппаратное обеспечение (Телефоны)";
                    oSheet.Cells[LastRow + 2, 1].Font.Bold = true;
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    oSheet.Cells[LastRow + 1, 1] = "Название";
                    oSheet.Cells[LastRow + 1, 2] = "Количество";
                    oSheet.Cells[LastRow + 1, 3] = "Сумма";
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                }
                else
                {
                    oSheet.Cells[LastRow, 1] = "Аппаратное обеспечение (Телефоны)";
                    oSheet.Cells[LastRow, 1].Font.Bold = true;
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    oSheet.Cells[LastRow + 1, 1] = "Название";
                    oSheet.Cells[LastRow + 1, 2] = "Количество";
                    oSheet.Cells[LastRow + 1, 3] = "Сумма";
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                }
            }    
            if (phone1Count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = Phone1_Text.Text;
                oSheet.Cells[LastRow + 1, 2] = phone1Count.Value;
                oSheet.Cells[LastRow + 1, 3] = phone1Count.Value * start_phonePrices[0];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (phone2Count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = Phone2_Text.Text;
                oSheet.Cells[LastRow + 1, 2] = phone2Count.Value;
                oSheet.Cells[LastRow + 1, 3] = phone2Count.Value * start_phonePrices[1];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (phone3Count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = Phone3_Text.Text;
                oSheet.Cells[LastRow + 1, 2] = phone3Count.Value;
                oSheet.Cells[LastRow + 1, 3] = phone3Count.Value * start_phonePrices[2];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (phone4Count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = Phone4_Text.Text;
                oSheet.Cells[LastRow + 1, 2] = phone4Count.Value;
                oSheet.Cells[LastRow + 1, 3] = phone4Count.Value * start_phonePrices[3];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (phone1Count.Value > 0 || phone2Count.Value > 0 || phone3Count.Value > 0 || phone4Count.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 2] = "Итого:";
                oSheet.Cells[LastRow + 1, 3] = phone1Count.Value * start_phonePrices[0] + phone2Count.Value * start_phonePrices[1] + phone3Count.Value * start_phonePrices[2] + phone4Count.Value * start_phonePrices[3];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (month1Price.Value > 0 || month2Price.Value > 2 || month3Price.Value > 0 || month4Price.Value > 0 || month5Price.Value>0)
            {

                if (LastRow - 1 != 0)
                {
                    oSheet.Cells[LastRow + 2, 1] = "Оплата телефонии";
                    oSheet.Cells[LastRow + 2, 1].Font.Bold = true;
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    oSheet.Cells[LastRow + 1, 1] = "Название";
                    oSheet.Cells[LastRow + 1, 2] = "Количество";
                    oSheet.Cells[LastRow + 1, 3] = "Сумма";
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                }
                else
                {
                    oSheet.Cells[LastRow, 1] = "Оплата телефонии";
                    oSheet.Cells[LastRow, 1].Font.Bold = true;
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    oSheet.Cells[LastRow + 1, 1] = "Название";
                    oSheet.Cells[LastRow + 1, 2] = "Количество";
                    oSheet.Cells[LastRow + 1, 3] = "Сумма";
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                }
            }        
            if (month1Price.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = monthpay1_text.Text;
                oSheet.Cells[LastRow + 1, 2] = month1Price.Value;
                oSheet.Cells[LastRow + 1, 3] = month1Price.Value * start_mothPayPrices[0];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (month2Price.Value > 0)
            {
                if (isCompanyPay_ChBox.Checked == true)
                {
                    oSheet.Cells[LastRow + 1, 1] = monthpay2_text.Text;
                    oSheet.Cells[LastRow + 1, 2] = month2Price.Value;
                    oSheet.Cells[LastRow + 1, 3] = month2Price.Value * start_mothPayPrices[1];
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                }
            }
            if (month3Price.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = monthpay3_text.Text;
                oSheet.Cells[LastRow + 1, 2] = month3Price.Value;
                oSheet.Cells[LastRow + 1, 3] = month3Price.Value * start_mothPayPrices[2];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (month4Price.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = monthpay4_text.Text;
                oSheet.Cells[LastRow + 1, 2] = month4Price.Value;
                oSheet.Cells[LastRow + 1, 3] = month4Price.Value * start_mothPayPrices[3];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (month5Price.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = monthpay5_text.Text;
                oSheet.Cells[LastRow + 1, 2] = month5Price.Value;
                oSheet.Cells[LastRow + 1, 3] = month5Price.Value * start_mothPayPrices[4];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (month1Price.Value > 0 || month2Price.Value > 2 || month3Price.Value > 0 || month4Price.Value > 0 || month5Price.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 2] = "Итого:";
                oSheet.Cells[LastRow + 1, 3] = month1Price.Value * start_mothPayPrices[0] + month2Price.Value * start_mothPayPrices[1] + month3Price.Value * start_mothPayPrices[2] + month4Price.Value * start_mothPayPrices[3] + month5Price.Value * start_mothPayPrices[4];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (pcChoosingPanelCount1.Value > 0 || pcChoosingPanelCount2.Value > 0 || pcChoosingPanelCount3.Value > 0 || pcChoosingPanelCount4.Value > 0 || pcChoosingPanelCount5.Value > 0 || pcChoosingPanelCount6.Value > 0)
            {
                if (LastRow - 1 != 0)
                {
                    oSheet.Cells[LastRow + 2, 1] = "Аппаратное обеспечение (Компьютеры)";
                    oSheet.Cells[LastRow + 2, 1].Font.Bold = true;
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    oSheet.Cells[LastRow + 1, 1] = "Название";
                    oSheet.Cells[LastRow + 1, 2] = "Количество";
                    oSheet.Cells[LastRow + 1, 3] = "Сумма";
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                }
                else
                {
                    oSheet.Cells[LastRow, 1] = "Аппаратное обеспечение (Компьютеры)";
                    oSheet.Cells[LastRow, 1].Font.Bold = true;
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    oSheet.Cells[LastRow+1, 1] = "Название";
                    oSheet.Cells[LastRow+1, 2] = "Количество";
                    oSheet.Cells[LastRow+1, 3] = "Сумма";
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                }
            }
            if (pcChoosingPanelCount1.Value > 0){
                oSheet.Cells[LastRow + 1, 1] = pcChoose1_Text.Text;
                oSheet.Cells[LastRow + 1, 2] = pcChoosingPanelCount1.Value;
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (pcChoosingPanelCount2.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = pcChoose2_Text.Text;
                oSheet.Cells[LastRow + 1, 2] = pcChoosingPanelCount2.Value;
                oSheet.Cells[LastRow + 1, 3] = pcChoosingPanelCount2.Value * start_pcPrices[0];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (pcChoosingPanelCount3.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = pcChoose3_Text.Text;
                oSheet.Cells[LastRow + 1, 2] = pcChoosingPanelCount3.Value;
                oSheet.Cells[LastRow + 1, 3] = pcChoosingPanelCount3.Value * start_pcPrices[1];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (pcChoosingPanelCount4.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = pcChoose4_Text.Text;
                oSheet.Cells[LastRow + 1, 2] = pcChoosingPanelCount4.Value;
                oSheet.Cells[LastRow + 1, 3] = pcChoosingPanelCount4.Value * start_pcPrices[2];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (pcChoosingPanelCount5.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = isTripleEncryptionNeed_ChBox.Text;
                oSheet.Cells[LastRow + 1, 2] = pcChoosingPanelCount5.Value;
                oSheet.Cells[LastRow + 1, 3] = pcChoosingPanelCount5.Value * start_pcPrices[3];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (pcChoosingPanelCount6.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = isProgrammBlockNeed_ChBox.Text;
                oSheet.Cells[LastRow + 1, 2] = pcChoosingPanelCount6.Value;
                oSheet.Cells[LastRow + 1, 3] = pcChoosingPanelCount6.Value * start_pcPrices[4];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (pcChoosingPanelCount1.Value > 0 || pcChoosingPanelCount2.Value > 0 || pcChoosingPanelCount3.Value > 0 || pcChoosingPanelCount4.Value > 0 || pcChoosingPanelCount5.Value > 0 || pcChoosingPanelCount6.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 2] = "Итого:";
               oSheet.Cells[LastRow + 1, 3] =  pcChoosingPanelCount2.Value * start_pcPrices[0] + pcChoosingPanelCount3.Value * start_pcPrices[1] + pcChoosingPanelCount4.Value * start_pcPrices[2] + pcChoosingPanelCount5.Value * start_pcPrices[3] + pcChoosingPanelCount6.Value * start_pcPrices[4];
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (rentServerCount1.Value > 0 || rentServerCount2.Value > 0)
            {
                if (LastRow - 1 != 0)
                {
                    oSheet.Cells[LastRow + 2, 1] = "Ежемесячные расходы на сервера и услуги";
                    oSheet.Cells[LastRow + 2, 1].Font.Bold = true;
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    oSheet.Cells[LastRow + 1, 1] = "Название";
                    oSheet.Cells[LastRow + 1, 2] = "Количество";
                    oSheet.Cells[LastRow + 1, 3] = "Сумма";
                    oSheet.Cells[LastRow + 1, 1].Font.Size = 12;
                    oSheet.Cells[LastRow + 1, 2].Font.Size = 12;
                    oSheet.Cells[LastRow + 1, 3].Font.Size = 12;
                    oSheet.Cells[LastRow + 1, 1].Font.Bold = true;
                    oSheet.Cells[LastRow + 1, 2].Font.Bold = true;
                    oSheet.Cells[LastRow + 1, 3].Font.Bold = true;
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                }
                else
                {
                    oSheet.Cells[LastRow, 1] = "Ежемесячные расходы на сервера и услуги";
                    oSheet.Cells[LastRow, 1].Font.Bold = true;
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    oSheet.Cells[LastRow + 1, 1] = "Название";
                    oSheet.Cells[LastRow + 1, 2] = "Количество";
                    oSheet.Cells[LastRow + 1, 3] = "Сумма";
                    oSheet.Cells[LastRow + 1, 1].Font.Size = 12;
                    oSheet.Cells[LastRow + 1, 2].Font.Size = 12;
                    oSheet.Cells[LastRow + 1, 3].Font.Size = 12;
                    oSheet.Cells[LastRow + 1, 1].Font.Bold = true;
                    oSheet.Cells[LastRow + 1, 2].Font.Bold = true;
                    oSheet.Cells[LastRow + 1, 3].Font.Bold = true;
                    LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                }

            }
            if (rentServerCount1.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = rentServer1_text.Text;
                oSheet.Cells[LastRow + 1, 2] = rentServerCount1.Value;
                oSheet.Cells[LastRow + 1, 3] = rentServerCount1.Value * start_serverRenPrices;
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (rentServerCount2.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 1] = rentServer2_Text.Text;
                oSheet.Cells[LastRow + 1, 2] = rentServerCount2.Value;
                oSheet.Cells[LastRow + 1, 3] = rentServerCount2.Value * start_serverRenPrices;
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
            if (rentServerCount1.Value > 0 || rentServerCount2.Value > 0)
            {
                oSheet.Cells[LastRow + 1, 2] = "Итого:";
                oSheet.Cells[LastRow + 1, 3] = rentServerCount1.Value * start_serverRenPrices + rentServerCount2.Value * start_serverRenPrices;
                LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }

            oSheet.Cells[LastRow+2, 2] = "Итого разовый платеж:";
            oSheet.Cells[LastRow + 2, 2].Font.Bold = true;
           var globaloverAll = oSheet.Cells[LastRow+2, 3] =  A1_count.Value * start_serverPrices[0] + A2_count.Value * start_serverPrices[1] + A3_count.Value * start_serverPrices[2] + A4_count.Value * start_serverPrices[3] + A5_count.Value * start_serverPrices[4] + A6_count.Value * start_serverPrices[5] + A1extra_count.Value * start_extraPrices[0] + A2extra_count.Value * start_extraPrices[1] + A3extra_count.Value * start_extraPrices[2] + A4extra_count.Value * start_extraPrices[3] + A5extra_count.Value * start_extraPrices[4] + A6extra_count.Value * start_extraPrices[5]+ phone1Count.Value * start_phonePrices[0] + phone2Count.Value * start_phonePrices[1] + phone3Count.Value * start_phonePrices[2] + phone4Count.Value * start_phonePrices[3]+ pcChoosingPanelCount2.Value * start_pcPrices[0] + pcChoosingPanelCount3.Value * start_pcPrices[1] + pcChoosingPanelCount4.Value * start_pcPrices[2] + pcChoosingPanelCount5.Value * start_pcPrices[3] + pcChoosingPanelCount6.Value * start_pcPrices[4]+ rentServerCount1.Value * start_serverRenPrices + rentServerCount2.Value * start_serverRenPrices;
            oSheet.Cells[LastRow + 2, 3].Font.Bold = true;
            oSheet.Cells[LastRow+3, 2] = "Итого ежемесячный платеж:";
            oSheet.Cells[LastRow + 3, 2].Font.Bold = true;
         var  monthoverAll = oSheet.Cells[LastRow + 3, 3] = month1Price.Value * start_mothPayPrices[0] + month2Price.Value * start_mothPayPrices[1] + month3Price.Value * start_mothPayPrices[2] + month4Price.Value * start_mothPayPrices[3] + month5Price.Value * start_mothPayPrices[4];
            oSheet.Cells[LastRow + 3, 3].Font.Bold = true;
            oSheet.Cells[LastRow+4, 2] = "Общий итог:";
            oSheet.Cells[LastRow + 4, 2].Font.Bold = true;
            oSheet.Cells[LastRow+4, 3] = globaloverAll+monthoverAll+"$";
            oSheet.Cells[LastRow + 4, 3].Font.Bold = true;
            LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            for (int i = LastRow; i >= 1; i--)
            {
                oSheet.StandardWidth = 30;
                var usingCells = oSheet.Range[
                    oSheet.Cells[i, 1],
                    oSheet.Cells[i, 3]];
                usingCells.Font.Size = 12;
                oSheet.Rows[i].HorizontalAlignment = -4108;
                oSheet.Rows[i].VerticalAlignment = -4108;
                string name = "Название";

                if ((i % 2) == 0)
                {
                    if (oSheet.Cells[i, 2].Value != null)
                    {
                        if (oSheet.Cells[i, 2].Text != "Итого:" && oSheet.Cells[i, 1].Text != "Серверная структура" && oSheet.Cells[i, 1].Text != "Серверная структура (дополнительные сервисы)" && oSheet.Cells[i, 1].Text != "Аппаратное обеспечение (Компьютеры)" && oSheet.Cells[i, 1].Text != "Ежемесячные расходы на сервера и услуги" && oSheet.Cells[i, 1].Text != "Аппаратное обеспечение (Телефоны)")
                        {
                            var columnHeadingsRange = oSheet.Range[
                                oSheet.Cells[i, 1],
                                oSheet.Cells[i, 3]];

                            columnHeadingsRange.Interior.Color = Color.FromArgb(0, 241, 242, 242);
                        }
                    }
                }
                else
                {
                    if (oSheet.Cells[i, 2].Value != null)
                    {
                        if (oSheet.Cells[i, 2].Text != "Итого:" && oSheet.Cells[i, 1].Text != "Серверная структура" && oSheet.Cells[i, 1].Text != "Серверная структура (дополнительные сервисы)" && oSheet.Cells[i, 1].Text != "Аппаратное обеспечение (Компьютеры)" && oSheet.Cells[i, 1].Text != "Ежемесячные расходы на сервера и услуги" && oSheet.Cells[i, 1].Text != "Аппаратное обеспечение (Телефоны)")
                        {
                            var columnHeadingsRange = oSheet.Range[
                            oSheet.Cells[i, 1],
                            oSheet.Cells[i, 3]];

                            columnHeadingsRange.Interior.Color = Color.FromArgb(0, 248, 249, 248);
                        }
                    }
                }
                string overprice = "Итого:";
                if (oSheet.Cells[i, 1].Text == name|| oSheet.Cells[i, 2].Text == overprice)
                {
                    Console.WriteLine(111111);
                    var LastInteriorRange = oSheet.Range[
                oSheet.Cells[i, 1],
                oSheet.Cells[i, 3]];
                    LastInteriorRange.Font.Size = 12;
                    LastInteriorRange.Font.Bold = true;
                    LastInteriorRange.Interior.Color = Color.FromArgb(0, 214, 209, 235);
                }
                /* if (oSheet.Cells[i, 2].Text == "0" || oSheet.Cells[i, 2].Value == null)
                 {
                     oSheet.Rows[i].Delete();
                     LastRow = oBook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                 }*/
            }
            /*var FirstInteriorRange = oSheet.Range[oSheet.Cells[1, 1],
                oSheet.Cells[1, 3]];
            FirstInteriorRange.Interior.Color = Color.FromArgb(0, 214, 209, 235);
            var LastInteriorRange = oSheet.Range[
                oSheet.Cells[LastRow + 1, 1],
                oSheet.Cells[LastRow + 1, 3]];
            LastInteriorRange.Interior.Color = Color.FromArgb(0, 214, 209, 235);*/
           // Console.WriteLine(LastRow);
            oSheet.Columns.AutoFit();
            oSheet.Rows.AutoFit();
            try
            {
                oBook.SaveAs(filePath);
                oBook.Close();
                oApp.Quit();
            }
            catch
            {
                string message = "Код ошибки ХХХ. Пожалуйста, обратитесь в службу технической поддержки!";
                string caption = "Ошибка!";
                MessageBoxButtons buttons = MessageBoxButtons.OK;
                DialogResult result;
                result = MessageBox.Show(message, caption, buttons);
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    oApp.DisplayAlerts = false;
                    oBook.Close();
                    oApp.Quit();

                    Application.Exit();
                }
            }
            Application.Exit();
            
        }

        private void A1_count_ValueChanged(object sender, EventArgs e)
        {

        }

        private void A2_count_ValueChanged(object sender, EventArgs e)
        {

        }

        private void A3_count_ValueChanged(object sender, EventArgs e)
        {

        }

        private void A4_count_ValueChanged(object sender, EventArgs e)
        {

        }

        private void A5_count_ValueChanged(object sender, EventArgs e)
        {

        }

        private void A6_count_ValueChanged(object sender, EventArgs e)
        {

        }

        private void A1extra_count_ValueChanged(object sender, EventArgs e)
        {
            if (A1extra_count.Value != 0 || A2extra_count.Value != 0 || A4extra_count.Value != 0 || A3extra_count.Value != 0 || A5extra_count.Value != 0 || A6extra_count.Value != 0)
                Aextra_doneButton.Enabled = true;
            else
                Aextra_doneButton.Enabled = false;
        }

        private void A2extra_count_ValueChanged(object sender, EventArgs e)
        {
            if (A1extra_count.Value != 0 || A2extra_count.Value != 0 || A4extra_count.Value != 0 || A3extra_count.Value != 0 || A5extra_count.Value != 0 || A6extra_count.Value != 0)
                Aextra_doneButton.Enabled = true;
            else
                Aextra_doneButton.Enabled = false;
        }

        private void A3extra_count_ValueChanged(object sender, EventArgs e)
        {
            if(A1extra_count.Value != 0 || A2extra_count.Value != 0 || A4extra_count.Value != 0 || A3extra_count.Value != 0 || A5extra_count.Value != 0 || A6extra_count.Value != 0)
                Aextra_doneButton.Enabled = true;
            else
                Aextra_doneButton.Enabled = false;
        }

        private void A4extra_count_ValueChanged(object sender, EventArgs e)
        {
            if (A1extra_count.Value != 0 || A2extra_count.Value != 0 || A4extra_count.Value != 0 || A3extra_count.Value != 0 || A5extra_count.Value != 0 || A6extra_count.Value != 0)
                Aextra_doneButton.Enabled = true;
            else
                Aextra_doneButton.Enabled = false;
        }

        private void A5extra_count_ValueChanged(object sender, EventArgs e)
        {
            if (A1extra_count.Value != 0 || A2extra_count.Value != 0 || A4extra_count.Value != 0 || A3extra_count.Value != 0 || A5extra_count.Value != 0 || A6extra_count.Value != 0)
                Aextra_doneButton.Enabled = true;
            else
                Aextra_doneButton.Enabled = false;
        }

        private void A6extra_count_ValueChanged(object sender, EventArgs e)
        {
            if (A1extra_count.Value != 0 || A2extra_count.Value != 0 || A4extra_count.Value != 0 || A3extra_count.Value != 0 || A5extra_count.Value != 0 || A6extra_count.Value != 0)
                Aextra_doneButton.Enabled = true;
            else
                Aextra_doneButton.Enabled = false;
        }

        private void Phone1Count_ValueChanged(object sender, EventArgs e)
        {
            if (phone1Count.Value != 0 || phone2Count.Value != 0 || phone3Count.Value != 0 || phone4Count.Value != 0)
                phoneDoneButton.Enabled = true;
            else
                phoneDoneButton.Enabled = false;
        }

        private void Phone2Count_ValueChanged(object sender, EventArgs e)
        {
            if (phone1Count.Value != 0 || phone2Count.Value != 0 || phone3Count.Value != 0 || phone4Count.Value != 0)
                phoneDoneButton.Enabled = true;
            else
                phoneDoneButton.Enabled = false;
        }

        private void Phone3Count_ValueChanged(object sender, EventArgs e)
        {
            if (phone1Count.Value != 0 || phone2Count.Value != 0 || phone3Count.Value != 0 || phone4Count.Value != 0)
                phoneDoneButton.Enabled = true;
            else
                phoneDoneButton.Enabled = false;
        }

        private void Phone4Count_ValueChanged(object sender, EventArgs e)
        {
            if (phone1Count.Value != 0 || phone2Count.Value != 0 || phone3Count.Value != 0 || phone4Count.Value != 0)
                phoneDoneButton.Enabled = true;
            else
                phoneDoneButton.Enabled = false;
        }

        private void PcChoosingPanelCount1_ValueChanged(object sender, EventArgs e)
        {
            pcChoosingPanelCount3.Value = pcChoosingPanelCount1.Value;
            pcChoosingPanelCount4.Value = pcChoosingPanelCount1.Value;
            pcChoosingPanelCount2.Maximum = pcChoosingPanelCount1.Value;
            pcChoosingPanelCount5.Maximum = pcChoosingPanelCount1.Value- pcChoosingPanelCount6.Value;
            pcChoosingPanelCount6.Maximum = pcChoosingPanelCount1.Value - pcChoosingPanelCount5.Value;
            if ( pcChoosingPanelCount1.Value != 0 || pcChoosingPanelCount2.Value != 0 || pcChoosingPanelCount3.Value != 0 || pcChoosingPanelCount4.Value != 0|| pcChoosingPanelCount5.Value != 0|| pcChoosingPanelCount6.Value != 0)
                pcChoosingDone_Button.Enabled = true;
            else
                pcChoosingDone_Button.Enabled = false;
        }

        private void PcChoosingPanelCount2_ValueChanged(object sender, EventArgs e)
        {

            if (pcChoosingPanelCount1.Value != 0 || pcChoosingPanelCount2.Value != 0 || pcChoosingPanelCount3.Value != 0 || pcChoosingPanelCount4.Value != 0 || pcChoosingPanelCount5.Value != 0 || pcChoosingPanelCount6.Value != 0)
                pcChoosingDone_Button.Enabled = true;
            else
                pcChoosingDone_Button.Enabled = false;

        }

        private void PcChoosingPanelCount3_ValueChanged(object sender, EventArgs e)
        {
            if (pcChoosingPanelCount1.Value != 0 || pcChoosingPanelCount2.Value != 0 || pcChoosingPanelCount3.Value != 0 || pcChoosingPanelCount4.Value != 0 || pcChoosingPanelCount5.Value != 0 || pcChoosingPanelCount6.Value != 0)
                pcChoosingDone_Button.Enabled = true;
            else
                pcChoosingDone_Button.Enabled = false;
        }

        private void PcChoosingPanelCount4_ValueChanged(object sender, EventArgs e)
        {
            if (pcChoosingPanelCount1.Value != 0 || pcChoosingPanelCount2.Value != 0 || pcChoosingPanelCount3.Value != 0 || pcChoosingPanelCount4.Value != 0 || pcChoosingPanelCount5.Value != 0 || pcChoosingPanelCount6.Value != 0)
                pcChoosingDone_Button.Enabled = true;
            else
                pcChoosingDone_Button.Enabled = false;
       
        }

        private void PcChoosingPanelCount5_ValueChanged(object sender, EventArgs e)
        {
            pcChoosingPanelCount5.Maximum = pcChoosingPanelCount1.Value - pcChoosingPanelCount6.Value;
            pcChoosingPanelCount6.Maximum = pcChoosingPanelCount1.Value - pcChoosingPanelCount5.Value;
            if (pcChoosingPanelCount1.Value != 0 || pcChoosingPanelCount2.Value != 0 || pcChoosingPanelCount3.Value != 0 || pcChoosingPanelCount4.Value != 0 || pcChoosingPanelCount5.Value != 0 || pcChoosingPanelCount6.Value != 0)
                pcChoosingDone_Button.Enabled = true;
            else
                pcChoosingDone_Button.Enabled = false;
        }

        private void PcChoosingPanelCount6_ValueChanged(object sender, EventArgs e)
        {
            pcChoosingPanelCount5.Maximum = pcChoosingPanelCount1.Value - pcChoosingPanelCount6.Value;
            pcChoosingPanelCount6.Maximum = pcChoosingPanelCount1.Value - pcChoosingPanelCount5.Value;
            if (pcChoosingPanelCount1.Value != 0 || pcChoosingPanelCount2.Value != 0 || pcChoosingPanelCount3.Value != 0 || pcChoosingPanelCount4.Value != 0 || pcChoosingPanelCount5.Value != 0 || pcChoosingPanelCount6.Value != 0)
                pcChoosingDone_Button.Enabled = true;
            else
                pcChoosingDone_Button.Enabled = false;
        }

        private void Month1Price_ValueChanged(object sender, EventArgs e)
        {
            if (month1Price.Value != 0 || month2Price.Value != 0 || month3Price.Value != 0 || month4Price.Value != 0 || month5Price.Value != 0)
                pcChoosingDone_Button.Enabled = true;
            else
                pcChoosingDone_Button.Enabled = false;
        }

        private void Month2Price_ValueChanged(object sender, EventArgs e)
        {

        }

        private void Month3Price_ValueChanged(object sender, EventArgs e)
        {
            month3Price.Maximum = (phone1Count.Value + phone2Count.Value + phone3Count.Value);
            if ((month4Price.Value + month3Price.Value) < (phone1Count.Value+phone2Count.Value+phone3Count.Value))
            {
                monthPhone_doneButton.Enabled = false;

            }
            else
            {
                monthPhone_doneButton.Enabled = true;

            }
        }

        private void Month4Price_ValueChanged(object sender, EventArgs e)
        {
            month4Price.Maximum = (phone1Count.Value + phone2Count.Value + phone3Count.Value);
            if ((month4Price.Value + month3Price.Value) < (phone1Count.Value + phone2Count.Value + phone3Count.Value))
            {
                monthPhone_doneButton.Enabled = false;

            }
            else
            {
                monthPhone_doneButton.Enabled = true;

            }
        }

        private void Month5Price_ValueChanged(object sender, EventArgs e)
        {

        }

        private void IsClientPay_ChBox_CheckedChanged(object sender, EventArgs e)
        {
            if(isClientPay_ChBox.Checked == true)
            {
                month1Price.Enabled = true;
                isCompanyPay_ChBox.Enabled = false;
            }
            else
            {
                month1Price.Enabled = false;
                isCompanyPay_ChBox.Enabled = true;
            }
        }

        private void IsCompanyPay_ChBox_CheckedChanged(object sender, EventArgs e)
        {
            if(isCompanyPay_ChBox.Checked == true)
            {

                month2Price.Enabled = true;
                isClientPay_ChBox.Enabled = false;
            }
            else
            {

                month2Price.Enabled = false;
                isClientPay_ChBox.Enabled = true;
            }
        }







        private void PhoneChoose_Text_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void PcChoose_Text_TextChanged(object sender, EventArgs e)
        {

        }

        private void ContextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void IsServiceNeed_CheckedChanged(object sender, EventArgs e)
        {
            if (isServiceNeed.Checked == true)
            {
                rentServerCount2.Enabled = true;
            }
            else
            {
                rentServerCount2.Enabled = false;
            }
        }

        private void IsTechnicalSupportNeed_ChBox_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void RentServer_Panel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void RentServerCount2_ValueChanged(object sender, EventArgs e)
        {
            int currentPrice = 0;
            if (rentServerCount2.Value == 0)
            {
                currentPrice =0;
                serverService_Text.Text = currentPrice.ToString() + "$";
            }
            if (rentServerCount2.Value > 0)
            {
                currentPrice = start_serverServicePrice;
                serverService_Text.Text = currentPrice.ToString() + "$";
            }
            if (rentServerCount2.Value > 5)
            {
                currentPrice = start_serverServicePrice * 2;
                serverService_Text.Text = currentPrice.ToString() + "$";
            }
            if (rentServerCount2.Value > 10)
            {
                currentPrice += start_serverServicePrice;
                serverService_Text.Text = currentPrice.ToString() + "$";
            }

            if (rentServerCount2.Value > 15)
            {
                currentPrice += start_serverServicePrice;
                serverService_Text.Text = currentPrice.ToString() + "$";
            }
            if (rentServerCount2.Value > 20)
            {
                currentPrice += start_serverServicePrice;
                serverService_Text.Text = currentPrice.ToString() + "$";
            }
            if (rentServerCount2.Value > 25)
            {
                currentPrice += start_serverServicePrice;
                serverService_Text.Text = currentPrice.ToString() + "$";
            }
        }
    }
}
