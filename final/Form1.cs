using GemBox.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;



namespace final
{
    public partial class Form1 : Form
    {
        //public static string SetValueForText1 = "";
        private bool isCollapsed;
        private bool isCollapsed1;
        string[,] S = new string[5, 40];


        public Form1()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown18_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown13_ValueChanged(object sender, EventArgs e)
        {

        }

        private void UPLOAD_Click(object sender, EventArgs e)
        {
            
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            var workbook = new ExcelFile();
            var worksheet = workbook.Worksheets.Add("fcsd");

            worksheet.Cells[0, 0].Value = "START(sec)";
            worksheet.Cells[0, 1].Value = "STOP(sec)";
            worksheet.Cells[1, 0].Value = ( Convert.ToInt32(U11.Text)*60)+Convert.ToInt32(U1.Text);
            worksheet.Cells[2, 0].Value = (Convert.ToInt32(U12.Text) * 60) + Convert.ToInt32(U2.Text);
            worksheet.Cells[3, 0].Value = (Convert.ToInt32(U13.Text) * 60) + Convert.ToInt32(U3.Text);
            worksheet.Cells[4, 0].Value = (Convert.ToInt32(U14.Text) * 60) + Convert.ToInt32(U4.Text);
            worksheet.Cells[5, 0].Value = (Convert.ToInt32(U15.Text) * 60) + Convert.ToInt32(U5.Text);
            worksheet.Cells[6, 0].Value = (Convert.ToInt32(U16.Text) * 60) + Convert.ToInt32(U6.Text);
            worksheet.Cells[7, 0].Value = (Convert.ToInt32(U17.Text) * 60) + Convert.ToInt32(U7.Text);
            worksheet.Cells[8, 0].Value = (Convert.ToInt32(U18.Text) * 60) + Convert.ToInt32(U8.Text);
            worksheet.Cells[9, 0].Value = (Convert.ToInt32(U19.Text) * 60) + Convert.ToInt32(U9.Text);
            worksheet.Cells[10,0].Value = (Convert.ToInt32(U110.Text) * 60) + Convert.ToInt32(U10.Text);
            worksheet.Cells[1, 1].Value = (Convert.ToInt32(D11.Text) * 60) + Convert.ToInt32(D1.Text);
            worksheet.Cells[2, 1].Value = (Convert.ToInt32(D12.Text) * 60) + Convert.ToInt32(D2.Text);
            worksheet.Cells[3, 1].Value = (Convert.ToInt32(D13.Text) * 60) + Convert.ToInt32(D3.Text);
            worksheet.Cells[4, 1].Value = (Convert.ToInt32(D14.Text) * 60) + Convert.ToInt32(D4.Text);
            worksheet.Cells[5, 1].Value = (Convert.ToInt32(D15.Text) * 60) + Convert.ToInt32(D5.Text);
            worksheet.Cells[6, 1].Value = (Convert.ToInt32(D16.Text) * 60) + Convert.ToInt32(D6.Text);
            worksheet.Cells[7, 1].Value = (Convert.ToInt32(D17.Text) * 60) + Convert.ToInt32(D7.Text);
            worksheet.Cells[8, 1].Value = (Convert.ToInt32(D18.Text) * 60) + Convert.ToInt32(D8.Text);
            worksheet.Cells[9, 1].Value = (Convert.ToInt32(D19.Text) * 60) + Convert.ToInt32(D9.Text);
            worksheet.Cells[10,1].Value = (Convert.ToInt32(D110.Text) * 60) + Convert.ToInt32(D10.Text);
            workbook.Save("yCSD.xlsx");
        }

        private void RESET_Click(object sender, EventArgs e)
        {
            U1.Text = "0";
            D1.Text = "0";
            U2.Text = "0";
            D2.Text = "0";
            U3.Text = "0";
            D3.Text = "0";
            U4.Text = "0";
            D4.Text = "0";
            U5.Text = "0";
            D5.Text = "0";
            U6.Text = "0";
            D6.Text = "0";
            U7.Text = "0";
            D7.Text = "0";
            U8.Text = "0";
            D8.Text = "0";
            U9.Text = "0";
            D9.Text = "0";
            U10.Text = "0";
            D10.Text = "0";
            U11.Text = "0";
            D11.Text = "0";
            U12.Text = "0";
            D12.Text = "0";
            U13.Text = "0";
            D13.Text = "0";
            U14.Text = "0";
            D14.Text = "0";
            U15.Text = "0";
            D15.Text = "0";
            U16.Text = "0";
            D16.Text = "0";
            U17.Text = "0";
            D17.Text = "0";
            U18.Text = "0";
            D18.Text = "0";
            U19.Text = "0";
            D19.Text = "0";
            U110.Text = "0";
            D110.Text = "0";
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if(isCollapsed)
            {
                panel.Height += 30;
                if(panel.Size==panel.MaximumSize)
                {
                    timer1.Stop();
                    isCollapsed = false;
                }
            }
            else
            {
                panel.Height -= 30;
                if (panel.Size == panel.MinimumSize)
                {
                    timer1.Stop();
                    isCollapsed = true;
                }
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (isCollapsed1)
            {
                panel1.Height += 30;
                if (panel1.Size == panel1.MaximumSize)
                {
                    timer2.Stop();
                    isCollapsed1 = false;
                }
            }
            else
            {
                panel1.Height -= 30;
                if (panel1.Size == panel1.MinimumSize)
                {
                    timer2.Stop();
                    isCollapsed1 = true;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            timer1.Start();
        }

        private void save1_Click(object sender, EventArgs e)
        {
            timer2.Start();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void slot1_Click(object sender, EventArgs e)
        {
            S[0,0] = U11.Text;
            S[0,1] = U1.Text;
            S[0,2] = D11.Text;
            S[0,3] = D1.Text;
            S[0,5] = U2.Text;
            S[0,7] = D2.Text;
            S[0,9] = U3.Text;
            S[0,11] = D3.Text;
            S[0,13] = U4.Text;
            S[0,15] = D4.Text;
            S[0,17] = U5.Text;
            S[0,19] = D5.Text;
            S[0,21] = U6.Text;
            S[0,23] = D6.Text;
            S[0,25] = U7.Text;
            S[0,27] = D7.Text;
            S[0,29] = U8.Text;
            S[0,31] = D8.Text;
            S[0,33] = U9.Text;
            S[0,35] = D9.Text;
            S[0,37] = U10.Text;
            S[0,39] = D10.Text;      
            S[0,4] = U12.Text;
            S[0,6] = D12.Text;
            S[0,8] = U13.Text;
            S[0,10] = D13.Text;
            S[0,12] = U14.Text;
            S[0,14] = D14.Text;
            S[0,16] = U15.Text;
            S[0,18] = D15.Text;
            S[0,20] = U16.Text;
            S[0,22] = D16.Text;
            S[0,24] = U17.Text;
            S[0,26] = D17.Text;
            S[0,28] = U18.Text;
            S[0,30] = D18.Text;
            S[0,32] = U19.Text;
            S[0,34] = D19.Text;
            S[0,36] = U110.Text;
            S[0,38] = D110.Text;
            slot1.Text = newname.Text;
            list1.Text = newname.Text;
            timer2.Start();
        }

        private void slot2_Click(object sender, EventArgs e)
        {
            S[1,1] = U1.Text;
            S[1,3] = D1.Text;
            S[1,5] = U2.Text;
            S[1,7] = D2.Text;
            S[1,9] = U3.Text;
            S[1,11] = D3.Text;
            S[1,13] = U4.Text;
            S[1,15] = D4.Text;
            S[1,17] = U5.Text;
            S[1,19] = D5.Text;
            S[1,21] = U6.Text;
            S[1,23] = D6.Text;
            S[1,25] = U7.Text;
            S[1,27] = D7.Text;
            S[1,29] = U8.Text;
            S[1,31] = D8.Text;
            S[1,33] = U9.Text;
            S[1,35] = D9.Text;
            S[1,37] = U10.Text;
            S[1,39] = D10.Text;
            S[1,0] = U11.Text;
            S[1,2] = D11.Text;
            S[1,4] = U12.Text;
            S[1,6] = D12.Text;
            S[1,8] = U13.Text;
            S[1,10] = D13.Text;
            S[1,12] = U14.Text;
            S[1,14] = D14.Text;
            S[1,16] = U15.Text;
            S[1,18] = D15.Text;
            S[1,20] = U16.Text;
            S[1,22] = D16.Text;
            S[1,24] = U17.Text;
            S[1,26] = D17.Text;
            S[1,28] = U18.Text;
            S[1,30] = D18.Text;
            S[1,32] = U19.Text;
            S[1,34] = D19.Text;
            S[1,36] = U110.Text;
            S[1,38] = D110.Text;
            slot2.Text = newname.Text;
            list2.Text = newname.Text;
            timer2.Start();
        }

        private void slot3_Click(object sender, EventArgs e)
        {
            S[2,1] = U1.Text;
            S[2,3] = D1.Text;
            S[2,5] = U2.Text;
            S[2,7] = D2.Text;
            S[2,9] = U3.Text;
            S[2,11] = D3.Text;
            S[2,13] = U4.Text;
            S[2,15] = D4.Text;
            S[2,17] = U5.Text;
            S[2,19] = D5.Text;
            S[2,21] = U6.Text;
            S[2,23] = D6.Text;
            S[2,25] = U7.Text;
            S[2,27] = D7.Text;
            S[2,29] = U8.Text;
            S[2,31] = D8.Text;
            S[2,33] = U9.Text;
            S[2,35] = D9.Text;
            S[2,37] = U10.Text;
            S[2,39] = D10.Text;
            S[2,0] = U11.Text;
            S[2,2] = D11.Text;
            S[2,4] = U12.Text;
            S[2,6] = D12.Text;
            S[2,8] = U13.Text;
            S[2,10] = D13.Text;
            S[2,12] = U14.Text;
            S[2,14] = D14.Text;
            S[2,16] = U15.Text;
            S[2,18] = D15.Text;
            S[2,20] = U16.Text;
            S[2,22] = D16.Text;
            S[2,24] = U17.Text;
            S[2,26] = D17.Text;
            S[2,28] = U18.Text;
            S[2,30] = D18.Text;
            S[2,32] = U19.Text;
            S[2,34] = D19.Text;
            S[2,36] = U110.Text;
            S[2,38] = D110.Text;
            slot3.Text = newname.Text;
            list3.Text = newname.Text;
            timer2.Start();
        }

        private void slot4_Click(object sender, EventArgs e)
        {
            S[3,1] = U1.Text;
            S[3,3] = D1.Text;
            S[3,5] = U2.Text;
            S[3,7] = D2.Text;
            S[3,9] = U3.Text;
            S[3,11] = D3.Text;
            S[3,13] = U4.Text;
            S[3,15] = D4.Text;
            S[3,17] = U5.Text;
            S[3,19] = D5.Text;
            S[3,21] = U6.Text;
            S[3,23] = D6.Text;
            S[3,25] = U7.Text;
            S[3,27] = D7.Text;
            S[3,29] = U8.Text;
            S[3,31] = D8.Text;
            S[3,33] = U9.Text;
            S[3,35] = D9.Text;
            S[3,37] = U10.Text;
            S[3,39] = D10.Text;
            S[3,0] = U11.Text;
            S[3,2] = D11.Text;
            S[3,4] = U12.Text;
            S[3,6] = D12.Text;
            S[3,8] = U13.Text;
            S[3,10] = D13.Text;
            S[3,12] = U14.Text;
            S[3,14] = D14.Text;
            S[3,16] = U15.Text;
            S[3,18] = D15.Text;
            S[3,20] = U16.Text;
            S[3,22] = D16.Text;
            S[3,24] = U17.Text;
            S[3,26] = D17.Text;
            S[3,28] = U18.Text;
            S[3,30] = D18.Text;
            S[3,32] = U19.Text;
            S[3,34] = D19.Text;
            S[3,36] = U110.Text;
            S[3,38] = D110.Text;
            slot4.Text = newname.Text;
            list4.Text = newname.Text;
            timer2.Start();
        }

        private void slot5_Click(object sender, EventArgs e)
        {
            S[4,1] = U1.Text;
            S[4,3] = D1.Text;
            S[4,5] = U2.Text;
            S[4,7] = D2.Text;
            S[4,9] = U3.Text;
            S[4,11] = D3.Text;
            S[4,13] = U4.Text;
            S[4,15] = D4.Text;
            S[4,17] = U5.Text;
            S[4,19] = D5.Text;
            S[4,21] = U6.Text;
            S[4,23] = D6.Text;
            S[4,25] = U7.Text;
            S[4,27] = D7.Text;
            S[4,29] = U8.Text;
            S[4,31] = D8.Text;
            S[4,33] = U9.Text;
            S[4,35] = D9.Text;
            S[4,37] = U10.Text;
            S[4,39] = D10.Text;
            S[4,0] = U11.Text;
            S[4,2] = D11.Text;
            S[4,4] = U12.Text;
            S[4,6] = D12.Text;
            S[4,8] = U13.Text;
            S[4,10] = D13.Text;
            S[4,12] = U14.Text;
            S[4,14] = D14.Text;
            S[4,16] = U15.Text;
            S[4,18] = D15.Text;
            S[4,20] = U16.Text;
            S[4,22] = D16.Text;
            S[4,24] = U17.Text;
            S[4,26] = D17.Text;
            S[4,28] = U18.Text;
            S[4,30] = D18.Text;
            S[4,32] = U19.Text;
            S[4,34] = D19.Text;
            S[4,36] = U110.Text;
            S[4,38] = D110.Text;
            slot5.Text = newname.Text;
            list5.Text = newname.Text;
            timer2.Start();
        }

        private void list1_Click(object sender, EventArgs e)
        {
            U1.Text = S[0,1];
            U2.Text = S[0,5];
            U3.Text = S[0,9];
            U4.Text = S[0,13];
            U5.Text = S[0,17];
            U6.Text = S[0,21];
            U7.Text = S[0,25];
            U8.Text = S[0,29];
            U9.Text = S[0,33];
            U10.Text = S[0,37];
            D1.Text = S[0,3];
            D2.Text = S[0,7];
            D3.Text = S[0,11];
            D4.Text = S[0,15];
            D5.Text = S[0,19];
            D6.Text = S[0,23];
            D7.Text = S[0,27];
            D8.Text = S[0,31];
            D9.Text = S[0,35];
            D10.Text = S[0,39];
            U11.Text = S[0,0];
            U12.Text = S[0,4];
            U13.Text = S[0,8];
            U14.Text = S[0,12];
            U15.Text = S[0,16];
            U16.Text = S[0,20];
            U17.Text = S[0,24];
            U18.Text = S[0,28];
            U19.Text = S[0,32];
            U110.Text = S[0,36];
            D11.Text = S[0,2];
            D12.Text = S[0,6];
            D13.Text = S[0,10];
            D14.Text = S[0,14];
            D15.Text = S[0,18];
            D16.Text = S[0,22];
            D17.Text = S[0,26];
            D18.Text = S[0,30];
            D19.Text = S[0,34];
            D110.Text = S[0,38];
            timer1.Start();
        }

        private void list2_Click(object sender, EventArgs e)
        {
            U1.Text = S[1,1];
            U2.Text = S[1,5];
            U3.Text = S[1,9];
            U4.Text = S[1,13];
            U5.Text = S[1,17];
            U6.Text = S[1,21];
            U7.Text = S[1,25];
            U8.Text = S[1,29];
            U9.Text = S[1,33];
            U10.Text = S[1,37];
            D1.Text = S[1,3];
            D2.Text = S[1,7];
            D3.Text = S[1,11];
            D4.Text = S[1,15];
            D5.Text = S[1,19];
            D6.Text = S[1,23];
            D7.Text = S[1,27];
            D8.Text = S[1,31];
            D9.Text = S[1,35];
            D10.Text = S[1,39];
            U11.Text = S[1,0];
            U12.Text = S[1,4];
            U13.Text = S[1,8];
            U14.Text = S[1,12];
            U15.Text = S[1,16];
            U16.Text = S[1,20];
            U17.Text = S[1,24];
            U18.Text = S[1,28];
            U19.Text = S[1,32];
            U110.Text = S[1,36];
            D11.Text = S[1,2];
            D12.Text = S[1,6];
            D13.Text = S[1,10];
            D14.Text = S[1,14];
            D15.Text = S[1,18];
            D16.Text = S[1,22];
            D17.Text = S[1,26];
            D18.Text = S[1,30];
            D19.Text = S[1,34];
            D110.Text = S[1,38];
            timer1.Start();
        }

        private void list3_Click(object sender, EventArgs e)
        {
            U1.Text = S[2,1];
            U2.Text = S[2,5];
            U3.Text = S[2,9];
            U4.Text = S[2,13];
            U5.Text = S[2,17];
            U6.Text = S[2,21];
            U7.Text = S[2,25];
            U8.Text = S[2,29];
            U9.Text = S[2,33];
            U10.Text = S[2,37];
            D1.Text = S[2,3];
            D2.Text = S[2,7];
            D3.Text = S[2,11];
            D4.Text = S[2,15];
            D5.Text = S[2,19];
            D6.Text = S[2,23];
            D7.Text = S[2,27];
            D8.Text = S[2,31];
            D9.Text = S[2,35];
            D10.Text = S[2,39];
            U11.Text = S[2,0];
            U12.Text = S[2,4];
            U13.Text = S[2,8];
            U14.Text = S[2,12];
            U15.Text = S[2,16];
            U16.Text = S[2,20];
            U17.Text = S[2,24];
            U18.Text = S[2,28];
            U19.Text = S[2,32];
            U110.Text = S[2,36];
            D11.Text = S[2,2];
            D12.Text = S[2,6];
            D13.Text = S[2,10];
            D14.Text = S[2,14];
            D15.Text = S[2,18];
            D16.Text = S[2,22];
            D17.Text = S[2,26];
            D18.Text = S[2,30];
            D19.Text = S[2,34];
            D110.Text = S[2,38];
            timer1.Start();
        }

        private void list4_Click(object sender, EventArgs e)
        {
            U1.Text = S[3,1];
            U2.Text = S[3,5];
            U3.Text = S[3,9];
            U4.Text = S[3,13];
            U5.Text = S[3,17];
            U6.Text = S[3,21];
            U7.Text = S[3,25];
            U8.Text = S[3,29];
            U9.Text = S[3,33];
            U10.Text = S[3,37];
            D1.Text = S[3,3];
            D2.Text = S[3,7];
            D3.Text = S[3,11];
            D4.Text = S[3,15];
            D5.Text = S[3,19];
            D6.Text = S[3,23];
            D7.Text = S[3,27];
            D8.Text = S[3,31];
            D9.Text = S[3,35];
            D10.Text = S[3,39];
            U11.Text = S[3,0];
            U12.Text = S[3,4];
            U13.Text = S[3,8];
            U14.Text = S[3,12];
            U15.Text = S[3,16];
            U16.Text = S[3,20];
            U17.Text = S[3,24];
            U18.Text = S[3,28];
            U19.Text = S[3,32];
            U110.Text = S[3,36];
            D11.Text = S[3,2];
            D12.Text = S[3,6];
            D13.Text = S[3,10];
            D14.Text = S[3,14];
            D15.Text = S[3,18];
            D16.Text = S[3,22];
            D17.Text = S[3,26];
            D18.Text = S[3,30];
            D19.Text = S[3,34];
            D110.Text = S[3,38];
            timer1.Start();
        }

        private void list5_Click(object sender, EventArgs e)
        {
            U1.Text = S[4,1];
            U2.Text = S[4,5];
            U3.Text = S[4,9];
            U4.Text = S[4,13];
            U5.Text = S[4,17];
            U6.Text = S[4,21];
            U7.Text = S[4,25];
            U8.Text = S[4,29];
            U9.Text = S[4,33];
            U10.Text = S[4,37];
            D1.Text = S[4,3];
            D2.Text = S[4,7];
            D3.Text = S[4,11];
            D4.Text = S[4,15];
            D5.Text = S[4,19];
            D6.Text = S[4,23];
            D7.Text = S[4,27];
            D8.Text = S[4,31];
            D9.Text = S[4,35];
            D10.Text = S[4,39];
            U11.Text = S[4,0];
            U12.Text = S[4,4];
            U13.Text = S[4,8];
            U14.Text = S[4,12];
            U15.Text = S[4,16];
            U16.Text = S[4,20];
            U17.Text = S[4,24];
            U18.Text = S[4,28];
            U19.Text = S[4,32];
            U110.Text = S[4,36];
            D11.Text = S[4,2];
            D12.Text = S[4,6];
            D13.Text = S[4,10];
            D14.Text = S[4,14];
            D15.Text = S[4,18];
            D16.Text = S[4,22];
            D17.Text = S[4,26];
            D18.Text = S[4,30];
            D19.Text = S[4,34];
            D110.Text = S[4,38];
            timer1.Start();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //this.Hide();
            Form f2 = new Form2();
            f2.Show();
        }

        private void text_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void label63_Click(object sender, EventArgs e)
        {
            label63.Text = Form2.SetValueForText1;
        }
        private void Form1_load(object sender, EventArgs e)
        {
            label64.Text = Form2.SetValueForText1;
        }
    }
}
