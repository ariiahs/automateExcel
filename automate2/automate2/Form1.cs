using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace automateExcel
{
    public partial class Form1 : Form
    {
        Excel.Application oXL;
        Excel._Workbook oWB;
        Excel._Worksheet oSheet;
        Excel.Range oRng;

        public Form1()
        {
            InitializeComponent();
        }
        private void openExcel(object sender, System.EventArgs e)
        {
           try
            {
                // Verify Launching of Excel
                oXL = new Excel.Application();
                oXL.Visible = true;

                //Verify creation of new WorkBook
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));

                //Verify creation of Worksheet
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                
                //Verify adding of new columns
                oSheet.Cells[1, 1] = "First Name";
                oSheet.Cells[1, 2] = "Last Name";
                oSheet.Cells[1, 3] = "Full Name";
                oSheet.Cells[1, 4] = "Monthly Salary";

                // Create an array to multiple values at once.
                string[,] empName = new string[4, 2];

                empName[0, 0] = "Adam";
                empName[0, 1] = "Cruz";
                empName[1, 0] = "Braedyn";
                empName[1, 1] = "Lee";
                empName[2, 0] = "Cherry";
                empName[2, 1] = "Tan";
                empName[3, 0] = "Diane";
                empName[3, 1] = "Jones";

                //Fill A2:B6 with an array of values (First and Last Names).
                oSheet.get_Range("A2", "B5").Value2 = empName;
                
                // Verify if Full name displays the FirstName and Last Name of a person
                oRng = oSheet.get_Range("C2", "C5");
                oRng.Formula = "=A2 & \" \" & B2";

                //Validate format for displayed values in monthly Salary
                oRng = oSheet.get_Range("D2", "D5");
                oRng.EntireColumn.AutoFit();
                oRng.NumberFormat = "$0.00";
                oSheet.Cells[2, 4] = "80000";
                oSheet.Cells[3, 4] = "40000";
                oSheet.Cells[4, 4] = "50000";
                oSheet.Cells[5, 4] = "60000";

                //Make sure Excel is visible and give the user control
                oXL.Visible = true;
                oXL.UserControl = true;

           }

                catch (Exception theException)
                {
                    String errorMessage;
                    errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, " Line: ");
                    errorMessage = String.Concat(errorMessage, theException.Source);

                    MessageBox.Show(errorMessage, "Error");
                }
        }

        private void insertColumn(object sender, EventArgs e)
        {
            //Verify inserting of new columns at the end of the record
            oSheet.Cells[1, 5] = "Annual Salary";
            oRng.NumberFormat = "$0.00";
        }

        private void cellFormat(object sender, EventArgs e)
        {
            //Format Columns
            oSheet.get_Range("A1", "E1").Font.Bold = true;
            oSheet.get_Range("A1", "E1").VerticalAlignment =
            Excel.XlVAlign.xlVAlignCenter;

            //AutoFit columns A:E.
            oRng = oSheet.get_Range("A1", "E1");
            oRng.EntireColumn.AutoFit();
        }

        private void insertFormula(object sender, EventArgs e)
        {
            //Verify total salary of each person for 1 year (monthly salary * 12 months)
            oRng = oSheet.get_Range("E2", "E5");
            oRng.Formula = "=D2 * 12";
            oRng.NumberFormat = "$0.00";
        }

        private void updateFormula(object sender, EventArgs e)
        {
            //Verify total salary when formula is updated to 2 years (monthly salary * 24 months)
            oRng = oSheet.get_Range("E2", "E5");
            oRng.Formula = "=D2 * 24";
            oRng.NumberFormat = "$0.00";

        }

        private void deleteColumn(object sender, EventArgs e)
        {
            //Verify if delete column button will delete the newly added Column
            oRng = oSheet.get_Range("E1", "E5");
            oRng.Delete();
        }

        private void deleteRow(object sender, EventArgs e)
        {
            //Verify if delete row button will delete a row
            oRng = oSheet.get_Range("A5", "E5");
            oRng.Delete();
        }

        private void closeExcel(object sender, EventArgs e)
        {
            //Verify closing of excel app 
            oXL.Quit();
            Application.Exit();

        }
  
    }
}   

            
       
           


