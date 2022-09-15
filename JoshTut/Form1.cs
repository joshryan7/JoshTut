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

namespace JoshTut
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int pos = textBox1.Text.IndexOf(" ");
            label1.Text = textBox1.Text.Substring(0, pos);
            int myL = textBox1.Text.Length;
            label2.Text = textBox1.Text.Substring(pos + 1, myL - pos - 1);
            label2.Text = label2.Text.ToUpper();

            if (label1.Text.ToUpper() == "JOSH")
            {
                label1.Text = "Sara";
            }
            if (label2.Text.ToUpper().Contains("KEL"))
            {
                label2.Text = "Parker";
            }
        }

        private void whileloopbtn_Click(object sender, EventArgs e)
        {
            //whileloopbtn
            //tb2
            tb2.Text = "";
            int start = Int32.Parse(textBox3.Text);
            int end = Int32.Parse(textBox4.Text);
            if (start < 1 || end > textBox1.Text.Length)
            {
                MessageBox.Show("My error message");
            }
            
            else
            {
                while (start < end)
                {

                    tb2.Text = tb2.Text + textBox1.Text.Substring(start - 1, 1);
                    start = start + 1;

                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //loopOutputlb
            int i = 1;
            while (i < 25)
            {
                i = i + 1;
                System.Threading.Thread.Sleep(100);
                loopOutputlb.Text = i.ToString();
                Application.DoEvents();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //characterToGrabtb
            //textbox2
            int grab = Int32.Parse(characterToGrabtb.Text);
            textBox2.Text = textBox1.Text.Substring(grab - 1, 1);
        }

        private void adddaysDoItbtn_Click(object sender, EventArgs e)
        {
            DateTime NewDate = addDaysDatePicker.Value.AddDays(Int32.Parse(adddaystb.Text));
            addDaysnewdatelb.Text = NewDate.ToShortDateString();


        }

        private void adddaystb_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                DateTime NewDate = addDaysDatePicker.Value.AddDays(Int32.Parse(adddaystb.Text));
                addDaysnewdatelb.Text = NewDate.ToShortDateString();
            }
            catch
            {

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void excelbtn_Click(object sender, EventArgs e)
        {

            JoshTut.dsInventory.SAM_Select_CurrentMachinesDataTable dt = new dsInventory.SAM_Select_CurrentMachinesDataTable();
            JoshTut.dsInventoryTableAdapters.SAM_Select_CurrentMachinesTableAdapter ta = new dsInventoryTableAdapters.SAM_Select_CurrentMachinesTableAdapter();

            dt = ta.GetData("CAT", false, false, false, DateTime.Parse("12/30/1899"), true);

            string cTemplateFile = @"u:\spike\CAT Current Inventory Report Template.xlsx";
            string cFileCopy = @"u:\spike\test.xlsx";
            int i = 1;

            #region create the file we are going to output to
            bool bOk = false;
            while (!bOk)
            {
                try
                {
                    System.IO.File.Copy(cTemplateFile, cFileCopy, true);
                    bOk = true;
                }
                catch (Exception e2)
                {

                    i++;
                    if (i > 50)
                    {
                        bOk = false;
                        MessageBox.Show("Error creating the excel file needed to then create the PDF.  Make sure the PDF is correct before you email it. Error: " + e2.Message);
                    }
                    cFileCopy = @"u:\spike\test_"+ (i.ToString()) + ".xlsx";
                }
            }
            #endregion
            char qt = '"';

            Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;

            Excel.Workbook wb = app.Workbooks.Open(cFileCopy, 0, false, 1, "", "", true,
             Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, false, false);

            Excel.Sheets wsheets = wb.Worksheets;

            Excel.Worksheet ws1 = (Microsoft.Office.Interop.Excel.Worksheet)wsheets.get_Item("Sheet1");

            int currentRow = 2;   // this needs to be the starting detail row

            Excel.Range dst;
            String lastouptput = "";

            // Example of how to fill in a cell
            //ws1.Cells[currentRow, 1] = "test";    // ,1 = Column A in teh Excel file,  2 = column B etc
            foreach (JoshTut.dsInventory.SAM_Select_CurrentMachinesRow r in dt)
            {
                // set the color of the 2 columns to Red  , columns S thru V in this example
                //dst = ws1.Range["S" + (currentRow).ToString() + ":V" + (currentRow).ToString()];
                //dst.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                if (r.SalesPlanStatus != lastouptput && lastouptput.Length > 0)
                {
                    currentRow++;
                }

                //if (r.SalesPlanStatus == "Won + Doc")
                //{

                //}

                if (currentRow == 2)
                {
                    dst = ws1.Range["E" + (currentRow).ToString() + ":E" + (currentRow).ToString()];
                    dst.NumberFormat = "[$$-en-US] #,##0.00";
                    //dst.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                }

                lastouptput = r.SalesPlanStatus;

                ws1.Cells[currentRow, 1] = r.invno;
                ws1.Cells[currentRow, 2] = r.jvpinvno;
                ws1.Cells[currentRow, 3] = r.descript;
                ws1.Cells[currentRow, 4] = r.machStatus;
                ws1.Cells[currentRow, 5] = r.SalesPlanStatus;
                ws1.Cells[currentRow, 6] = r.LastSalesPlanDate;
                ws1.Cells[currentRow, 7] = r.OriginalSalesPlanReturnAmount;
                ws1.Cells[currentRow, 8] = r.ProjectedReturn;
                ws1.Cells[currentRow, 9] = r.LastSalesPlanReturnAmount;
                ws1.Cells[currentRow, 10] = r.retail;

                ws1.Cells[currentRow, 12] = r.numberofOffers;
                ws1.Cells[currentRow, 13] = r.LastOfferAmount;
                ws1.Cells[currentRow, 14] = r.LatestOfferDate;
                ws1.Cells[currentRow, 15] = r.LatestOfferNote;
                ws1.Cells[currentRow, 16] = r.location;
                ws1.Cells[currentRow, 17] = r.loccity;

     
                currentRow++;   // go to the next row
            }


            wb.SaveAs(cFileCopy, Type.Missing, Type.Missing, Type.Missing,
              false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
              Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //wb.Close();
            
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void sortbtn_Click(object sender, EventArgs e)
        {
            List<string> Lnames = new List<string>();

            Lnames.Add("Sara");
            Lnames.Add("Josh");
            Lnames.Add("Matt");
            Lnames.Add("Liz");
            Lnames.Add("Andy");
            Lnames.Add("Denise");

            int i = 0;
            bool b = true;
            string temp = "";

            while (b)
            {
                b = false;
                i = 0;
                while (i < Lnames.Count()-1)
                {
                    if (Lnames[i].CompareTo(Lnames[i+1]) > 0)
                    {
                        temp = Lnames[i + 1];
                        Lnames[i + 1] = Lnames[i];
                        Lnames[i] = temp;
                        b = true;
                    }

                    i = i + 1;
                }

            }


            i = 2;


        }
    }
}
