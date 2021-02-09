using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ReadExcelFileApp
{
    public class ExcelFile
    {
        private string excelFilePath = string.Empty;
        private int rowNumber = 1; // define first row number to enter data in excel

        //ToT_WTParts Columns
        private string PARTNUMBER = "D";
        private string PARTNAME = "C";
        private string REVISION = "L";
        private string DEFAULTUNIT = "U";
        private string CREATED = "Q";
        private string MODIFYTIMESTAMP = "S";

        //ToT_BOM Columns



        Excel.Application myExcelApplication;
        Excel.Workbook myExcelWorkbook;
        Excel.Worksheet myExcelWorkSheet;

        public string ExcelFilePath
        {
            get { return excelFilePath; }
            set { excelFilePath = value; }
        }

        public int Rownumber
        {
            get { return rowNumber; }
            set { rowNumber = value; }
        }

        public void openExcel(string excelFilePath)
        {
            myExcelApplication = null;

            myExcelApplication = new Excel.Application(); // create Excel App
            myExcelApplication.DisplayAlerts = false; // turn off alerts


            myExcelWorkbook = (Excel.Workbook)(myExcelApplication.Workbooks._Open(excelFilePath, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
               System.Reflection.Missing.Value, System.Reflection.Missing.Value)); // open the existing excel file

            int numberOfWorkbooks = myExcelApplication.Workbooks.Count; // get number of workbooks (optional)

            myExcelWorkSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets["Ilox Part"]; // define in which worksheet, do you want to add data
            
            int numberOfSheets = myExcelWorkbook.Worksheets.Count; // get number of worksheets (optional)
        }

        public void addPDMDataToExcel(object[] itemArray)
        {
            var used = myExcelWorkSheet.UsedRange;

            string partNumber = itemArray[0].ToString();
            string title = itemArray[1].ToString();
            string rev = itemArray[2].ToString();
            string baseUnit = itemArray[3].ToString();
            string dateModified = itemArray[4].ToString();

            Excel.Range line = (Excel.Range)myExcelWorkSheet.Rows[Rownumber + 1];
            line.Insert();
                       
            myExcelWorkSheet.Cells[Rownumber, PARTNUMBER] = partNumber;
            myExcelWorkSheet.Cells[Rownumber, PARTNAME] = title;
            myExcelWorkSheet.Cells[Rownumber, REVISION] = rev;
            myExcelWorkSheet.Cells[Rownumber, DEFAULTUNIT] = baseUnit;
            myExcelWorkSheet.Cells[Rownumber, CREATED] = dateModified;
            myExcelWorkSheet.Cells[Rownumber, MODIFYTIMESTAMP] = dateModified;


            Rownumber++;  // if you put this method inside a loop, you should increase rownumber by one or wat ever is your logic

        }

        public void addOracleDataToExcel(object[] itemArray)
        {
            string description = itemArray[0].ToString();
            string partNumber = itemArray[1].ToString();
            string planningMakeBuyCode = itemArray[2].ToString();
            string revision = itemArray[3].ToString();
            string revisionEffectivityDate = itemArray[4].ToString();
            string primaryUOMCode = itemArray[5].ToString();

            Excel.Range line = (Excel.Range)myExcelWorkSheet.Rows[Rownumber + 1];
            line.Insert();

            

            myExcelWorkSheet.Cells[Rownumber, PARTNUMBER] = partNumber;
            myExcelWorkSheet.Cells[Rownumber, PARTNAME] = description;
            myExcelWorkSheet.Cells[Rownumber, REVISION] = revision;
            myExcelWorkSheet.Cells[Rownumber, DEFAULTUNIT] = primaryUOMCode;
            myExcelWorkSheet.Cells[Rownumber, CREATED] = revisionEffectivityDate;
            myExcelWorkSheet.Cells[Rownumber, MODIFYTIMESTAMP] = revisionEffectivityDate;

            Rownumber++;  // if you put this method inside a loop, you should increase rownumber by one or wat ever is your logic

        }

        public void closeExcel()
        {
            try
            {
                myExcelWorkbook.SaveAs(ExcelFilePath, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value); // Save data in excel


                myExcelWorkbook.Close(true, excelFilePath, System.Reflection.Missing.Value); // close the worksheet


            }
            finally
            {
                if (myExcelApplication != null)
                {
                    myExcelApplication.Quit(); // close the excel application
                }
            }

        }

    }
}
