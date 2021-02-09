using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;
using TextBox = System.Windows.Forms.TextBox;

namespace ReadExcelFileApp
{
    public partial class Form1 : Form
    {
        List<string> UnProcessedFiles = new List<string>();
        string cbSourceFilePath;
        string oracleSourceFilePath;
        string masterDestinationFilePath;
        string bomDestinationFilePath;
        string totPartsDestinationFilePath;
        string totBomnDestinationFilePath;
        private const string ILOX_PART_SHEET = "Ilox Part";
        private const string BOM_SHEET = "BOM";


        public Form1()
        {
            InitializeComponent();
            cbSourceFilePath = this.textBox1.Text;
            oracleSourceFilePath = this.textBox2.Text;
            masterDestinationFilePath = this.textBox3.Text;
            bomDestinationFilePath = this.textBox4.Text;
            totPartsDestinationFilePath = this.textBox5.Text;
            totBomnDestinationFilePath = this.textBox6.Text;
        }

        private void btnExport_Master_List_Of_Unique_Items(object sender, EventArgs e)
        {
            if (!(string.IsNullOrEmpty(cbSourceFilePath) || string.IsNullOrEmpty(oracleSourceFilePath) || string.IsNullOrEmpty(masterDestinationFilePath)))
            {
                string fileExt = string.Empty;
                string fileDateModified = string.Empty;

                string columns = "[Part Number],[Title],[REV],[Base Unit]";

                string[] bomsToParse = Directory.GetFiles(cbSourceFilePath, "*.*", SearchOption.AllDirectories);

                string result = string.Empty;
                DataTable dtExcel = new DataTable();
                foreach (string bom in bomsToParse)
                {
                    result = result + bom + System.Environment.NewLine;

                    string filePath = bom;
                    fileExt = Path.GetExtension(bom);
                    fileDateModified = File.GetLastWriteTime(filePath).ToString();
                    string sheetName = $"{BOM_SHEET}$";

                    if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                    {
                        try
                        {
                            if (dtExcel.Rows.Count == 0)
                                dtExcel = ReadExcel(filePath, fileExt, columns, sheetName, fileDateModified);
                            else
                                dtExcel.Merge(ReadExcel(filePath, fileExt, columns, sheetName, fileDateModified));

                            AddTopLevelRow(dtExcel, filePath, fileDateModified);
                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show(ex.Message.ToString());
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                    }
                }

                ExportExcel(dtExcel, masterDestinationFilePath, GetCurrentFileName(masterDestinationFilePath, "MasterListOfUniqueItems.xlsx"), "Unique Items");
                //ExportMasterListOfUniqueItems(dtExcel);

                string oracleSheet = "'Items Orgs; All Items$'";
                //string oracleFilePath = @"C:\Users\mruiz\Documents\PLM\BOM A_B_2014_2020\Oracle\Item Master and Latest Revision - All DPE orgs 11_18_2020.xlsx";
                string oracleColumns = "[DESCRIPTION],[Part Number],[PLANNING_MAKE_BUY_CODE],[REVISION],[Revision Effectivity Date],[PRIMARY_UOM_CODE]";
                DataTable dtOracle = ReadExcel(oracleSourceFilePath, Path.GetExtension(oracleSourceFilePath), oracleColumns, oracleSheet);

                RenameHeaders(dtExcel, columns + ",[Date modified]", "Parts");
                RenameHeaders(dtOracle, oracleColumns, "Parts");

                string totFileNameTemp = "ToT_WTPart_v1_original.xlsx";
                string totFileName = Path.Combine(totPartsDestinationFilePath, GetCurrentFileName(totPartsDestinationFilePath, $"ToT_WTPart_v1.xlsx"));
                File.Copy(Path.Combine(totPartsDestinationFilePath, totFileNameTemp), totFileName);

                InsertDataToExcel(dtOracle, totFileName, "'Ilox Part$'", "O");
                InsertDataToExcel(dtExcel, totFileName, "'Ilox Part$'", "C/B");

                //show files that could not be processed
                string issues = string.Empty;
                if (UnProcessedFiles.Count > 0)
                {
                    foreach (string file in UnProcessedFiles)
                    {
                        issues = issues + file + System.Environment.NewLine;
                    }
                    MessageBox.Show(issues);
                }
                MessageBox.Show("DONE");
            }
            else
            {
                MessageBox.Show("A path is missing.");
            }

        }

        private static string NO_DESCRIPTION = "NO DESCRIPTION";
        private void AddNoteToBlankField(DataTable dtExcel, string field)
        {
            if (dtExcel.Columns.Contains(field))
            {
                foreach (DataRow row in dtExcel.Rows)
                {
                    if (string.IsNullOrEmpty(row[field].ToString()))
                    {
                        row[field] = NO_DESCRIPTION;
                    }
                }
            }
        }

        private void RenameHeaders(DataTable dtExcel, string columns, string reportType)
        {
            string[] aCols = string.Join("", columns.ToCharArray().Where(c => !c.Equals('[') && !c.Equals(']')).ToList()).Split(',');

            Dictionary<string, string> colMapping = new Dictionary<string, string>();
            if (reportType.Equals("Parts"))
                AddPartWTMapping(colMapping);
            else
                AddBOMMapping(colMapping);

            foreach (string col in aCols)
            {
                if (colMapping.ContainsKey(col))
                {
                    dtExcel.Columns[col].ColumnName = colMapping[col];
                }
                else if (!string.IsNullOrEmpty(col))
                    dtExcel.Columns.Remove(col);
            }
        }

        public void AddPartWTMapping(Dictionary<string, string> dict)
        {
            dict.Add("Part Number", "partNumber");
            dict.Add("Title", "partName");
            dict.Add("REV", "revision");
            dict.Add("Base Unit", "defaultUnit");
            dict.Add("Date modified", "Created");
            dict.Add("DESCRIPTION", "partName");
            dict.Add("PLANNING_MAKE_BUY_CODE", "source");
            dict.Add("REVISION", "revision");
            dict.Add("Revision Effectivity Date", "Created");
            dict.Add("PRIMARY_UOM_CODE", "defaultUnit");
        }

        public void AddBOMMapping(Dictionary<string, string> dict)
        {
            dict.Add("Parent", "ASSEMBLYPARTNUMBER");
            dict.Add("Part Number", "ASSEMBLYPARTNUMBER");
            dict.Add("REV", "ASSEMBLYPARTVERSION");
            dict.Add("Child", "CONSTITUENTPARTNUMBER");
            dict.Add("QTY", "CONSTITUENTPARTQTY");
            dict.Add("Base Unit", "CONSTITUENTPARTUNIT");
            dict.Add("ASSEMBLY", "ASSEMBLYPARTNUMBER");
            dict.Add("REVISION", "ASSEMBLYPARTVERSION");
            dict.Add("COMPONENT", "CONSTITUENTPARTNUMBER");
            dict.Add("COMPONENT_QUANTITY", "CONSTITUENTPARTQTY");
            dict.Add("UOM", "CONSTITUENTPARTUNIT");
            dict.Add("ITEM_NUM", "LINENUMBER");
        }

        private void AddTopLevelRow(DataTable dtExcel, string filePath, string fileDateModified)
        {
            DataRow dr = dtExcel.NewRow();
            dr["Part Number"] = Path.GetFileNameWithoutExtension(filePath).Substring(0, 11);
            dr["Title"] = "MAIN ASSEMBLY";
            dr["REV"] = Path.GetFileNameWithoutExtension(filePath).Substring(15, 2);
            dr["Base Unit"] = "Each";
            dr["Date modified"] = fileDateModified;

            dtExcel.Rows.Add(dr);
        }

        private void btnExport_Parent_Child_Excel(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            string sheetName = $"{BOM_SHEET}$";
            string columns = "[Item],[QTY],[Part Number],[Component Type],[BOM Structure]";

            string[] bomsToParse = Directory.GetFiles(cbSourceFilePath, "*.*", SearchOption.AllDirectories);

            DataTable dtOutput = new DataTable();
            foreach (string bom in bomsToParse)
            {
                filePath = bom;
                fileExt = Path.GetExtension(bom);

                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel = new DataTable();
                        dtExcel = ReadExcel(filePath, fileExt, columns, sheetName);

                        //convert datatable to 2d array to write to excel
                        object[,] dtArray = new Object[dtExcel.Rows.Count + 1, dtExcel.Columns.Count];
                        dtArray[0, 0] = "0";
                        dtArray[0, 1] = string.Empty;
                        dtArray[0, 2] = Path.GetFileNameWithoutExtension(filePath).Substring(0, 11);
                        dtArray[0, 3] = string.Empty;
                        dtArray[0, 4] = string.Empty;
                        for (int row = 0; row < dtExcel.Rows.Count; row++)
                        {
                            for (int col = 0; col < dtExcel.Columns.Count; col++)
                            {
                                dtArray[row + 1, col] = dtExcel.Rows[row].ItemArray[col].ToString();
                            }
                        }
                        if (dtOutput.Rows.Count == 0)
                            dtOutput = GetParentChildData(dtArray, dtExcel.Rows.Count + 1);
                        else
                            dtOutput.Merge(GetParentChildData(dtArray, dtExcel.Rows.Count + 1));
                        dtArray = null;
                    }
                    catch (Exception ex)
                    {
                        // MessageBox.Show("File Issue Reported for:" + filePath + "\n" + ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }

            ExportExcel(RemovedDuplicatesDataTable(dtOutput), bomDestinationFilePath, GetCurrentFileName(bomDestinationFilePath, "BOMStructure.xlsx"), "BOM Structure");

            string bomFilePath = GetLatestFile(bomDestinationFilePath);
            string bomSheet = "'BOM Structure$'";
            string masterFilePath = GetLatestFile(masterDestinationFilePath);
            string masterSheet = "'Unique Items$'";
            string totFileNameTemp = "ToT_BOM_v1_original.xlsx";

            string totFileName = Path.Combine(totBomnDestinationFilePath, GetCurrentFileName(totBomnDestinationFilePath, $"ToT_BOM_v1.xlsx"));
            File.Copy(Path.Combine(totBomnDestinationFilePath, totFileNameTemp), totFileName);

            DataTable dtJoined = new DataTable();
            DataTable dtJoined2 = new DataTable();

            dtJoined = InnerJoin(bomFilePath, masterFilePath, bomSheet, masterSheet, "[Parent],[Child],[QTY],[REV]", "[Part Number],[REV]", "[Parent]", "[Part Number]");
            ExportExcel(RemovedDuplicatesDataTable(dtJoined), @"C:\Users\mruiz\Documents\PLM\R2\Output", "TempJoin.xlsx", "TEMP");


            dtJoined2 = InnerJoin(@"C:\Users\mruiz\Documents\PLM\R2\Output\TempJoin.xlsx", masterFilePath, "TEMP$", masterSheet, "[Parent],[Child],[QTY],[REV],[Base Unit]", "[Part Number],[Base Unit]", "[Child]", "[Part Number]");
            RenameHeaders(dtJoined2, "[Parent],[Child],[QTY],[REV],[Base Unit]", "BOM");
            dtJoined2 = dtJoined2.AsEnumerable().GroupBy(r => new { ASSEMBLYPARTNUMBER = r.Field<string>("ASSEMBLYPARTNUMBER"), CONSTITUENTPARTNUMBER = r.Field<string>("CONSTITUENTPARTNUMBER") }).Select(g => g.First()).CopyToDataTable();

            RemovedUOM(dtJoined2, "CONSTITUENTPARTQTY");

            InsertDataToExcel(dtJoined2, totFileName, $"{BOM_SHEET}$", "C/B", false);

            //******collect oracle data******
            string oracleFilePath = oracleSourceFilePath;
            //string oracleSheet = "205$";
            string[] oracleSheets = { "'209 BOMs$'", "'MST BOMs$'", "'211 BOMs$'", "'207 BOMs$'", "'306 BOMs$'", "'208 BOMs$'", "'305 BOMs$'", "'205 BOMs$'" };
            string oracleColumns = "[ASSEMBLY],[REVISION],[COMPONENT],[COMPONENT_QUANTITY],[UOM],[ITEM_NUM]";

            foreach (string sheet in oracleSheets)
            {
                DataTable dtOracle = ReadExcel(oracleFilePath, Path.GetExtension(oracleFilePath), oracleColumns, sheet);

                RenameHeaders(dtOracle, oracleColumns, "BOM");
                InsertDataToExcel(dtOracle, totFileName, $"{BOM_SHEET}$", "O", false);
            }


            //show files that could not be processed
            string issues = string.Empty;
            if (UnProcessedFiles.Count != 0)
            {
                foreach (string file in UnProcessedFiles)
                {
                    issues = issues + file + System.Environment.NewLine;
                }
                MessageBox.Show(issues);
            }
            MessageBox.Show("DONE");
        }

        private string GetLatestFile(string masterDestinationFilePath)
        {
            var directory = new DirectoryInfo(masterDestinationFilePath);

            var myFile = directory.GetFiles()
                         .OrderByDescending(f => f.LastWriteTime)
                         .First();

            return myFile.FullName;
        }

        private void RemovedUOM(DataTable dtJoined2, string columnToClean)
        {
            foreach (DataRow row in dtJoined2.Rows)
            {
                string value = row["CONSTITUENTPARTQTY"].ToString();
                double number = 0;
                if (!Double.TryParse(value, out number))
                {
                    row["CONSTITUENTPARTQTY"] = Regex.Replace(value, "[^0-9.]", "");
                }
            }
        }

        public DataTable GetParentChildData(object[,] allValues, int totalRows)
        {
            List<object> family = new List<object>();
            //creating excel to be saved

            //set first row with top parent, "level 0"
            DataTable dt = new DataTable();
            dt.Columns.Add("Parent Item");
            dt.Columns.Add("Parent");
            dt.Columns.Add("Child Item");
            dt.Columns.Add("Child");
            dt.Columns.Add("QTY");
            dt.Columns.Add("Component Type");
            dt.Columns.Add("BOM Structure");
            DataRow dr = null;



            for (int i = 0; i < totalRows; i++)
            {
                string parent_ItemNumber = allValues[i, 0].ToString();
                string parent_PartNumber = allValues[i, 2].ToString();

                // For example, if assembly's ItemNumber (first column) is 3.17.11.1, then it's children's itemNumbers would be 3.17.11.1.x, 3.17.11.1.x.x, ...
                // loop through the rows following the assembly, find all such children and add them to the delible list
                int count = allValues[i, 0].ToString().Count(x => x == '.');
                for (int j = i + 1; j < totalRows; j++)
                {
                    string[] itemNumberParts = allValues[j, 0].ToString().Split('.');
                    var parent = string.Join(".", itemNumberParts.Take(count + 1));

                    string child_ItemNumber = allValues[j, 0].ToString();
                    string child_PartNumber = allValues[j, 2].ToString();

                    if (itemNumberParts.Count() == 1 && !child_ItemNumber.Contains(".") && parent_ItemNumber == "0")
                    {
                        dr = dt.NewRow();
                        dr["Parent Item"] = allValues[i, 0].ToString();
                        dr["Parent"] = allValues[i, 2].ToString();
                        dr["Child Item"] = allValues[j, 0].ToString();
                        dr["Child"] = allValues[j, 2].ToString();
                        dr["QTY"] = allValues[j, 1].ToString();
                        dr["Component Type"] = allValues[j, 3].ToString();
                        dr["BOM Structure"] = allValues[j, 4].ToString();
                        dt.Rows.Add(dr);
                    }
                    else if (itemNumberParts.Count() > 1 && string.Join(".", itemNumberParts.Take(itemNumberParts.Count() - 1)).Equals(allValues[i, 0].ToString()))
                    {
                        //family.Add(allValues[j, 0].ToString());
                        family.Add(new { parent_ItemNumber, parent_PartNumber, child_ItemNumber, child_PartNumber });

                        //add row to datatable to allow writing to excel
                        dr = dt.NewRow();
                        dr["Parent Item"] = parent_ItemNumber;
                        dr["Parent"] = parent_PartNumber;
                        dr["Child Item"] = child_ItemNumber;
                        dr["Child"] = child_PartNumber;
                        dr["QTY"] = allValues[j, 1].ToString();
                        dr["Component Type"] = allValues[j, 3].ToString();
                        dr["BOM Structure"] = allValues[j, 4].ToString();
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        public DataTable RemovedDuplicatesDataTable(DataTable dtExcel)
        {
            DataTable dtExcelCopy = new DataTable();

            dtExcelCopy = dtExcel.AsEnumerable().GroupBy(r => new { Parent = r.Field<string>("Parent"), Child = r.Field<string>("Child") }).Select(g => g.First()).CopyToDataTable();

            return dtExcelCopy;
        }

        private void ExportExcel(DataTable dtExcel, string saveLocation, string fileName, string sheetName)
        {
            //filter for only unique part numbers
            //var distinctValues = dtExcel.AsEnumerable().GroupBy(r => r.Field<string>("Part Number")).Select(group => group.First()).CopyToDataTable();
            dataGridView1.Visible = true;
            dataGridView1.DataSource = dtExcel;

            //write to excel file
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dtExcel, sheetName);
            wb.SaveAs(Path.Combine(saveLocation, fileName));
        }

        private string GetCurrentFileName(string saveLocation, string fileName)
        {
            int fileCount = 0;

            if (!fileName.Contains("Temp"))
            {
                do
                {
                    fileCount++;
                }
                while
            (
                File.Exists(
                    Path.Combine(saveLocation, Path.GetFileNameWithoutExtension(fileName) + (fileCount > 0 ? "(" + fileCount.ToString() + ")" : "") + Path.GetExtension(fileName)))
                    );

                fileName = Path.GetFileNameWithoutExtension(fileName) + (fileCount > 0 ? "(" + fileCount.ToString() + ")" : "") + Path.GetExtension(fileName);
            }

            return fileName;
        }

        public void InsertDataToExcel(DataTable dtExcel, string fileName, string sheet, string fromSource, bool addDate = true)
        {
            string sql = null;
            int counter = 0;

            try
            {
                OleDbConnection MyConnection;
                OleDbCommand myCommand = new OleDbCommand();

                MyConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES';");
                MyConnection.Open();
                myCommand.Connection = MyConnection;

                string columnNames = "";
                foreach (DataColumn col in dtExcel.Columns)
                {
                    columnNames = columnNames + "[" + col.ColumnName + "],";
                }

                if (addDate)
                {
                    columnNames = columnNames + "[modifyTimestamp],";
                }
                if (!string.IsNullOrEmpty(fromSource))
                {
                    columnNames = columnNames + "[FROM_SOURCE]";
                }
                columnNames = columnNames.TrimEnd(',');

                //insert each row
                foreach (DataRow row in dtExcel.Rows)
                {
                    string rowValues = "";
                    foreach (var item in row.ItemArray)
                    {
                        if (item.GetType() != typeof(DBNull))
                        {
                            if (item.ToString().Contains("'"))
                            {
                                rowValues = rowValues + "'" + item.ToString().Replace("'", "''") + "',";
                            }
                            else
                                rowValues = rowValues + "'" + item + "',";
                        }
                        else
                            rowValues = rowValues + "null,";
                    }

                    if (addDate && dtExcel.Columns.Contains("Created"))
                    {
                        rowValues = rowValues + "'" + row["Created"].ToString() + "',";
                    }
                    if (!string.IsNullOrEmpty(fromSource))
                    {
                        rowValues = rowValues + $"'{fromSource}'";
                    }
                    rowValues = rowValues.TrimEnd(',');
                    sql = $"Insert into [{sheet}] ({columnNames}) values({rowValues});";
                    myCommand.CommandText = sql;
                    myCommand.ExecuteNonQuery();
                    counter++;
                }
                MyConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString() + "ROW: " + counter.ToString() + sql);
            }
        }

        public DataTable InnerJoin(string fileName1, string fileName2, string sheet1, string sheet2, string columns1, string columns2, string key1, string key2)
        {
            DataTable dtexcel = new DataTable();

            OleDbConnection oConn = null;
            OleDbCommand oComm = null;
            OleDbDataReader oRdr = null;
            string sConnString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName1 + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1';";
            string sCommand = String.Empty;
            try
            {
                oConn = new OleDbConnection(sConnString);
                oConn.Open();
                //sCommand = $"SELECT X.[CONSTITUENTPARTNUMBER], Y.[EBS_ORGANIZATION_CODE] FROM [BOM$] AS X INNER JOIN (SELECT [Part Number], [EBS_ORGANIZATION_CODE] FROM ['All Orgs; Diff Rvision$'] IN '{_dataExportFilePath}' 'Excel 12.0;HDR=YES;IMEX=1') AS Y ON X.[CONSTITUENTPARTNUMBER] = Y.[Part Number]";

                sCommand = $"SELECT {columns1} FROM [{sheet1}] AS X INNER JOIN (SELECT {columns2} FROM [{sheet2}] IN '{fileName2}' 'Excel 12.0;HDR=YES;IMEX=1') AS Y ON X.{key1} = Y.{key2}";
                oComm = new OleDbCommand(sCommand, oConn);
                oRdr = oComm.ExecuteReader();
                dtexcel.Load(oRdr);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (oRdr != null) oRdr.Close();
                oRdr = null;
                if (dtexcel != null) dtexcel.Dispose();
                oComm.Dispose();
                oConn.Close();
                oConn.Dispose();
            }
            return dtexcel;
        }

        public DataTable ReadExcel(string fileName, string fileExt, string columns, string sheet, string dateModified = null)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();

            if (fileExt.CompareTo(".xls") == 0)
            {
                //for below excel 2007 
                conn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + @";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1';";
            }
            else
            {
                //for above excel 2007
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1';";
            }
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt;
                    oleAdpt = new OleDbDataAdapter($"select {columns} from [{sheet}]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  

                    //For troubleshooting
                    //string[] missing = { "DPP-0224329", "DPP-0245136", "DPP-0784285", "DPP-0784347", "DPP-0790929", "DPP-0791676", "DPP-0817011", "DPP-0817226", "DPP-0842416", "DPP-0842652", "DPP-0845541", "DPP-0853562", "DPP-0853673", "DPP-0867893", "DPP-0867961", "DPP-0884692", "DPP-0892001", "DPP-0918363", "DPP-0926020", "DPP-0926196", "DPP-0927884", "DPP-0927954", "DPP-0938120", "DPP-0955376", "DPP-0958650", "DPP-1274532", "DPP-1286398", "DPP-1512373", "DPP-1512379" };
                    //string str = "";

                    //bool contains = false;
                    //foreach (string part in missing)
                    //{
                    //    contains = dtexcel.AsEnumerable().Any(r => part == r.Field<string>("Part Number"));
                    //    if (contains)
                    //    {
                    //        var found = dtexcel.AsEnumerable().Where(r => r.Field<string>("Part Number").Equals(part));
                    //        string missingName = found.Select(p => p.Field<string>("Title")).First();
                    //        string missingRev = found.Select(p => p.Field<string>("REV")).First();
                    //        str = str + $"PartNumber: {part} found in FileName: {fileName} has PartName:{missingName} REV:{missingRev}" +  System.Environment.NewLine;
                    //    }
                    //}
                    //string path = @"C:\Users\mruiz\Documents\PLM\R2\Output\Master List\MissingNameAndRevs.txt";
                    //if (!string.IsNullOrEmpty(str))
                    //{
                    //    using (StreamWriter sw = File.AppendText(path))
                    //    {
                    //        sw.WriteLine(str);
                    //    }
                    //}
                }
                catch (Exception ex)
                {
                    UnProcessedFiles.Add(fileName);
                    throw;
                }
            }

            if (!string.IsNullOrWhiteSpace(dateModified) && dtexcel.Columns.Count != 0)
            {
                dtexcel.Columns.Add("Date modified");
                foreach (DataRow dr in dtexcel.Rows)
                {
                    dr["Date modified"] = dateModified;
                }
            }
            return dtexcel;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close(); //to close the window(Form1)  
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog browser = new FolderBrowserDialog();
            browser.SelectedPath = @"C:\Users\mruiz\Documents\PLM";

            if (browser.ShowDialog() == DialogResult.OK)
            {
                cbSourceFilePath = browser.SelectedPath;
                textBox1.Text = cbSourceFilePath;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog browser = new OpenFileDialog();
            browser.InitialDirectory = @"C:\Users\mruiz\Documents\PLM";

            if (browser.ShowDialog() == DialogResult.OK)
            {
                oracleSourceFilePath = browser.SafeFileName;
                textBox2.Text = oracleSourceFilePath;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog browser = new FolderBrowserDialog();
            browser.SelectedPath = @"C:\Users\mruiz\Documents\PLM";

            if (browser.ShowDialog() == DialogResult.OK)
            {
                masterDestinationFilePath = browser.SelectedPath;
                textBox3.Text = masterDestinationFilePath;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {

        }

        //Map Supplier
        private void button10_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();

            dt = ReadExcel(oracleSourceFilePath, Path.GetExtension(oracleSourceFilePath), "[MANUFACTURER_NAME]", "'Mfg Names$'");
            dt.Columns["MANUFACTURER_NAME"].ColumnName = "Name";

            InsertDataToExcel(dt, Path.Combine(totPartsDestinationFilePath, GetLatestFile(totPartsDestinationFilePath)), "Supplier$", null, false); //GetCurrentFileName(totPartsDestinationFilePath, $"ToT_WTPart_v1.xlsx"))
        }

        //Map Manufacturer Part
        private void button11_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();

            dt = ReadExcel(oracleSourceFilePath, Path.GetExtension(oracleSourceFilePath), "[Part],[Manufacturer],[DESCRIPTION]", "'Mfg Part #s$'");
            dt.Columns["Part"].ColumnName = "partNumber";
            dt.Columns["DESCRIPTION"].ColumnName = "partName";
            dt.Columns["Manufacturer"].ColumnName = "manufacturerName";

            InsertDataToExcel(dt, Path.Combine(totPartsDestinationFilePath, GetLatestFile(totPartsDestinationFilePath)), "'Manufacturer Part$'", null, false); //GetCurrentFileName(totPartsDestinationFilePath, $"ToT_WTPart_v1.xlsx"))
        }

        //Map AML
        private void button12_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();

            dt = ReadExcel(oracleSourceFilePath, Path.GetExtension(oracleSourceFilePath), "[Item],[Part],[Manufacturer]", "'Mfg Part #s$'");
            dt.Columns["Item"].ColumnName = "OEMPartNumber";                 //Mfg Parts #s [G]
            //dt.Columns[""].ColumnName = "OEMVersion";                      //Items Orgs; All Items sheet REVISION[O]
            dt.Columns["Part"].ColumnName = "manufacturerPartNumber";        //Mfg Parts #s [D]
            dt.Columns["Manufacturer"].ColumnName = "manufacturerName";      //Mfg Parts #s [B]

            InsertDataToExcel(dt, Path.Combine(totPartsDestinationFilePath, GetLatestFile(totPartsDestinationFilePath)), "AML$", null, false);
        }

        //RegEx
        public DataTable CrossReferenceDPP(DataTable dataTable)
        {

            int maxStarNumbers = dataTable.Select().Max(r => r["Comments"].ToString().Count(c => c == '*'));

            //create required number of Star PN columns, since comments may contain more than one
            int col = 1;
            while (col <= maxStarNumbers)
            {
                dataTable.Columns.Add("Star PN" + col);
                col++;
            }

            foreach (DataRow row in dataTable.Rows)
            {
                string starNumber = row["Comments"].ToString();
                int starCount = starNumber.Count(c => c == '*');
                int count = 1;

                if (Regex.IsMatch(starNumber, "DPE\\*\\d{7,8}"))
                {
                    var numbers = Regex.Matches(starNumber, "DPE\\*\\d{7,8}").Cast<Match>().Select(m => m.Value);

                    foreach (string num in numbers)
                    {
                        row["Star PN" + count] = num;
                        count++;
                    }
                }
                //DPE ETO*12345678
                if (Regex.IsMatch(starNumber, "DPE\\sETO\\*\\d{7,8}"))
                {
                    var numbers = Regex.Matches(starNumber, "DPE\\sETO\\*\\d{7,8}").Cast<Match>().Select(m => m.Value);

                    foreach (string num in numbers)
                    {
                        row["Star PN" + count] = num;
                        count++;
                    }
                }
                if (Regex.IsMatch(starNumber, "DPE\\sETO\\sCTRLS\\*\\d{7,8}"))
                {
                    var numbers = Regex.Matches(starNumber, "DPE\\sETO\\sCTRLS\\*\\d{7,8}").Cast<Match>().Select(m => m.Value);

                    foreach (string num in numbers)
                    {
                        row["Star PN" + count] = num;
                        count++;
                    }
                }
                //DPE PARTS*12345678
                if (Regex.IsMatch(starNumber, "DPE\\sPARTS\\*\\d{7,8}"))
                {
                    var numbers = Regex.Matches(starNumber, "DPE\\sPARTS\\*\\d{7,8}").Cast<Match>().Select(m => m.Value);

                    foreach (string num in numbers)
                    {
                        row["Star PN" + count] = num;
                        count++;
                    }
                }
                //DPE SPARES*12345678
                if (Regex.IsMatch(starNumber, "DPE\\sSPARES\\*\\d{7,8}"))
                {
                    var numbers = Regex.Matches(starNumber, "DPE\\sSPARES\\*\\d{7,8}").Cast<Match>().Select(m => m.Value);

                    foreach (string num in numbers)
                    {
                        row["Star PN" + count] = num;
                        count++;
                    }
                }
                //DPE SYSTEM*12345678
                if (Regex.IsMatch(starNumber, "DPE\\sSYSTEM\\*\\d{7,8}"))
                {
                    var numbers = Regex.Matches(starNumber, "DPE\\sSYSTEM\\*\\d{7,8}").Cast<Match>().Select(m => m.Value);

                    foreach (string num in numbers)
                    {
                        row["Star PN" + count] = num;
                        count++;
                    }
                }
                //POSTAL DPE*12345678
                if (Regex.IsMatch(starNumber, "POSTAL\\sDPE\\s\\*\\d{7,8}"))
                {
                    var numbers = Regex.Matches(starNumber, "POSTAL\\sDPE\\s\\*\\d{7,8}").Cast<Match>().Select(m => m.Value);

                    foreach (string num in numbers)
                    {
                        row["Star PN" + count] = num;
                        count++;
                    }
                }
                //POSTAL CTRLS*
                if (Regex.IsMatch(starNumber, "POSTAL\\sCTRLS\\s\\*\\d{7,8}"))
                {
                    var numbers = Regex.Matches(starNumber, "POSTAL\\sCTRLS\\s\\*\\d{7,8}").Cast<Match>().Select(m => m.Value);

                    foreach (string num in numbers)
                    {
                        row["Star PN" + count] = num;
                        count++;
                    }
                }
                //DPE CTRLS*12345678
                if (Regex.IsMatch(starNumber, "DPE\\sCTRLS\\*\\d{7,8}"))
                {
                    var numbers = Regex.Matches(starNumber, "DPE\\sCTRLS\\*\\d{7,8}").Cast<Match>().Select(m => m.Value);

                    foreach (string num in numbers)
                    {
                        row["Star PN" + count] = num;
                        count++;
                    }
                }
            }
            return dataTable;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            string ecrSheet = @"C:\Users\mruiz\Documents\PLM\R2\Master_Cross_Reference_011821.xlsx";
            DataTable dt = ReadExcel(ecrSheet, Path.GetExtension(ecrSheet), "[Name],[Comments]", "SN_with_Star$");

            dt = CrossReferenceDPP(dt);
            string sql = null;
            int counter = 0;

            try
            {
                OleDbConnection MyConnection;
                OleDbCommand myCommand = new OleDbCommand();

                MyConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ecrSheet + ";Extended Properties='Excel 12.0;HDR=YES';");
                MyConnection.Open();
                myCommand.Connection = MyConnection;

                foreach (DataRow row in dt.Rows)
                {
                    sql = $"Update [SN_with_Star$] " +
                        $"Set [Star PN1]='{row["Star PN1"].ToString()}'" +
                        (!string.IsNullOrEmpty(row["Star PN2"].ToString()) ? $", [Star PN2]='{row["Star PN2"].ToString()}'" : "") +
                        (!string.IsNullOrEmpty(row["Star PN3"].ToString()) ? $", [Star PN3]='{row["Star PN3"].ToString()}'" : "") +
                        (!string.IsNullOrEmpty(row["Star PN4"].ToString()) ? $", [Star PN4] = '{row["Star PN4"].ToString()}'" : "") +
                        (!string.IsNullOrEmpty(row["Star PN5"].ToString()) ? $", [Star PN5]='{row["Star PN5"].ToString()}'" : "") +
                        $" where Name='{row["Name"].ToString()}';";

                    myCommand.CommandText = sql;
                    myCommand.ExecuteNonQuery();
                }
                counter++;
                MyConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString() + "ROW: " + counter.ToString() + sql);
            }
        }


        private string _filePath;
        private string _bomFilePath;
        private string _dataExportFilePath;
        private void button15_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDlg = new OpenFileDialog();

            // Show the FolderBrowserDialog.  
            DialogResult result = fileDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox7.Text = fileDlg.FileName;
                _filePath = fileDlg.FileName;
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDlg = new OpenFileDialog();

            // Show the FolderBrowserDialog.  
            DialogResult result = fileDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox8.Text = fileDlg.FileName;
                _bomFilePath = fileDlg.FileName;
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDlg = new OpenFileDialog();

            // Show the FolderBrowserDialog.  
            DialogResult result = fileDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox9.Text = fileDlg.FileName;
                _dataExportFilePath = fileDlg.FileName;
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            _filePath = ((TextBox)sender).Text;
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            _bomFilePath = ((TextBox)sender).Text;
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            _dataExportFilePath = ((TextBox)sender).Text;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            //get data from excel
            //DataTable dt = InnerJoin(_dataExportFilePath, _bomFilePath, "'Items Orgs; All Items$'", "BOM$",
            //     "[Part Number], [EBS_ORGANIZATION_CODE]", "[ASSEMBLYPARTNUMBER]", "[Part Number]", "[ASSEMBLYPARTNUMBER]");

            DataTable dt = new DataTable();

            OleDbConnection oConn = null;
            OleDbCommand oComm = null;
            OleDbDataReader oRdr = null;
            string sConnString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _bomFilePath + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1';";
            string sCommand = String.Empty;
            try
            {
                oConn = new OleDbConnection(sConnString);
                oConn.Open();
                sCommand = $"SELECT X.[ASSEMBLYPARTNUMBER], Y.[EBS_ORGANIZATION_CODE] FROM [{BOM_SHEET}$] AS X INNER JOIN (SELECT [Part Number], [EBS_ORGANIZATION_CODE] FROM ['All Orgs; Diff Rvision$'] IN '{_dataExportFilePath}' 'Excel 12.0;HDR=YES;IMEX=1') AS Y ON X.[ASSEMBLYPARTNUMBER] =Y.[Part Number]";
                oComm = new OleDbCommand(sCommand, oConn);
                oRdr = oComm.ExecuteReader();
                dt.Load(oRdr);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (oRdr != null) oRdr.Close();
                oRdr = null;
                if (dt != null) dt.Dispose();
                oComm.Dispose();
                oConn.Close();
                oConn.Dispose();
            }

            dt.Columns["ASSEMBLYPARTNUMBER"].ColumnName = "csvobjectNumber";
            dt.Columns["EBS_ORGANIZATION_CODE"].ColumnName = "csvtargetNumber";

            InsertDataToExcel(dt, _filePath, "ReleaseHistory-WTPart$", "", false);
            InsertDataToExcel(dt, _filePath, "ReleaseHistory-BOMHeader$", "", false);

            dt = new DataTable();

            oConn = null;
            oComm = null;
            oRdr = null;
            sConnString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _bomFilePath + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1';";
            sCommand = String.Empty;
            try
            {
                oConn = new OleDbConnection(sConnString);
                oConn.Open();
                sCommand = $"SELECT X.[CONSTITUENTPARTNUMBER], Y.[EBS_ORGANIZATION_CODE] FROM [{BOM_SHEET}$] AS X INNER JOIN (SELECT [Part Number], [EBS_ORGANIZATION_CODE] FROM ['All Orgs; Diff Rvision$'] IN '{_dataExportFilePath}' 'Excel 12.0;HDR=YES;IMEX=1') AS Y ON X.[CONSTITUENTPARTNUMBER] = Y.[Part Number]";
                oComm = new OleDbCommand(sCommand, oConn);
                oRdr = oComm.ExecuteReader();
                dt.Load(oRdr);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (oRdr != null) oRdr.Close();
                oRdr = null;
                if (dt != null) dt.Dispose();
                oComm.Dispose();
                oConn.Close();
                oConn.Dispose();
            }

            dt.Columns["CONSTITUENTPARTNUMBER"].ColumnName = "csvobjectNumber";
            dt.Columns["EBS_ORGANIZATION_CODE"].ColumnName = "csvtargetNumber";

            InsertDataToExcel(dt, _filePath, "ReleaseHistory-WTPart$", "", false);
        }

        public void TotWtPartsCleanUp(string filePath)
        {
            string copyFilePath = filePath;
            int count = 0;
            while (File.Exists(copyFilePath))
            {
                count++;
                copyFilePath = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + $"-Copy" + ((count > 0) ? count.ToString() : "") + "." + Path.GetExtension(filePath));
            }
            File.Copy(filePath, copyFilePath);

            DataTable dt = new DataTable();
            dt = ReadExcel(copyFilePath, Path.GetExtension(copyFilePath), "*", $"'{ILOX_PART_SHEET}$'");

            DataTable sortedDataTable = dt.AsEnumerable().GroupBy(r => r.Field<string>("partNumber")).OrderByDescending(g => g.Key).Select(x => x.First()).CopyToDataTable();

            AddNoteToBlankField(sortedDataTable, "partName");

            //make sure copy of excel rows are deleted
            string sql = null;
            int counter = 0;
            try
            {
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false;

                Excel.Workbook workbook = excelApp.Workbooks.Open(copyFilePath,
                     Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                Excel.Worksheet ws = new Worksheet();
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name.Equals(ILOX_PART_SHEET))
                    {
                        ws = worksheet;
                        break;
                    }
                }

                int cols = ws.UsedRange.Columns.Count;
                int rows = ws.UsedRange.Rows.Count;

                Excel.Range c1 = ws.Cells[2, 1];
                Excel.Range c2 = ws.Cells[rows, cols];
                // Range range = (Range)ws.get_Range("A2", "Z310593");

                var range = (Range)ws.get_Range(c1, c2);
                range.Delete(XlDeleteShiftDirection.xlShiftUp);
                workbook.Save();
                workbook.Close();
                excelApp.Quit();

                InsertDataToExcel(sortedDataTable, copyFilePath, $"'{ILOX_PART_SHEET}$'", null, false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        const string RAW_EXPORT_DATA_SUMMARY = @"C:\Users\mruiz\Downloads\Raw_Data_Export_Summary_012921.xlsx";
        private void UpdateRegion(string fileToUpdate)
        {
            //Replace "C/B" with its respected region => A=AsiaPac B=Brazil
            const string ALL_COLUMNS = "*";
            const string COMBINED_SHEET = "Combined";
            string sheetName = fileToUpdate.Contains("ToT_WTPart") ? ("'" + ILOX_PART_SHEET + "$'") : BOM_SHEET + "$";
            DataTable rawExportDataTable = ReadExcel(RAW_EXPORT_DATA_SUMMARY, Path.GetExtension(RAW_EXPORT_DATA_SUMMARY), ALL_COLUMNS, $"{COMBINED_SHEET}$");

            string sql = null;
            int counter = 0;
            try
            {
                OleDbConnection MyConnection;
                OleDbCommand myCommand = new OleDbCommand();

                MyConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileToUpdate + ";Extended Properties='Excel 12.0;HDR=YES';");
                MyConnection.Open();
                myCommand.Connection = MyConnection;

                foreach (DataRow row in rawExportDataTable.Rows)
                {
                    sql = $"Update [{sheetName}] " +
                        $"Set [FROM_SOURCE]='{row["region"].ToString()}'" +
                        //$" where partNumber='{row["partNumber"].ToString()}';";
                        $" where " + (fileToUpdate.Contains("ToT_WTPart") ? "partNumber" : "ASSEMBLYPARTNUMBER") + $"='{row["partNumber"].ToString()}';";

                    myCommand.CommandText = sql;
                    myCommand.ExecuteNonQuery();
                }
                counter++;
                MyConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString() + "ROW: " + counter.ToString() + sql);
            }
        }
        public void ReplaceChinaBrazilValue(string excel)
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false;

                Excel.Workbook workbook = excelApp.Workbooks.Open(excel,
                     Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                Excel.Worksheet ws = new Worksheet();
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name.Equals(excel.Contains("ToT_WTPart") ? ILOX_PART_SHEET : BOM_SHEET))
                    {
                        ws = worksheet;
                        break;
                    }
                }

                ws.Cells.Replace("C/B", "A");

                workbook.Save();
                workbook.Close();
                excelApp.Quit();

            }
            catch (Exception ex) { }
        }

        private string[] GetRegionPartNumbers(DataTable dataTable, string field, string region)
        {
            if (dataTable.Columns.Contains("partNumber") && dataTable.Columns.Contains(field))
            {
                return dataTable.AsEnumerable().Where(r => r.Field<string>(field).Equals(region)).Select(c => c.Field<string>("partNumber")).ToArray(); ;
            }
            return null;
        }

        private string totPartsPath;
        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            totPartsPath = ((TextBox)sender).Text;
        }

        public string GetFilePath()
        {
            OpenFileDialog fileDlg = new OpenFileDialog();
            string filePath = string.Empty;
            // Show the FolderBrowserDialog.  
            DialogResult result = fileDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                filePath = fileDlg.FileName;
            }
            return filePath;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            totPartsPath = GetFilePath();
            textBox10.Text = totPartsPath;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            TotWtPartsCleanUp(totPartsPath);
        }

        private void button20_Click(object sender, EventArgs e)
        {
            UpdateRegion(fileToUpdate);
        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        string fileToUpdate;
        private string cadPartReportFile;
        private string totPartsFile;
        private string totEPMBuildRuleFile;

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            fileToUpdate = ((TextBox)sender).Text;
        }

        private void button21_Click(object sender, EventArgs e)
        {
            fileToUpdate = GetFilePath();
            textBox11.Text = fileToUpdate;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            ReplaceChinaBrazilValue(fileToUpdate);
        }

        private void button23_Click(object sender, EventArgs e)
        {
            ReplaceMissingRevision(fileToUpdate);
        }

        //WT_Parts has 107 parts without revision
        private void ReplaceMissingRevision(string fileToUpdate)
        {
            string EPDM_OUTPUT_EXCEL = @"C:\Users\mruiz\Downloads\ePDM_Output.xlsx";
            string conn = string.Empty;

            DataTable revTable = ReadExcel(EPDM_OUTPUT_EXCEL, Path.GetExtension(EPDM_OUTPUT_EXCEL), "[Number], [Revision]", "tmp12$");
            DataTable updateTable = new DataTable();

            conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileToUpdate + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1';";
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt;
                    string sql = string.Empty;
                    OleDbConnection MyConnection;
                    OleDbCommand myCommand = new OleDbCommand();

                    oleAdpt = new OleDbDataAdapter($"select [partNumber],[revision] from ['{ILOX_PART_SHEET}$'] where revision is null", con); //here we read data from sheet1  
                    oleAdpt.Fill(updateTable); //fill excel data into dataTable  

                    foreach (DataRow row in updateTable.Rows)
                    {
                        DataRow[] foundRows = revTable.Select("Number like '%" + row.Field<string>("partNumber") + "%'");

                        if (foundRows.Count() > 0)
                        {
                            foreach (DataRow fr in foundRows)
                            {
                                if (fr.Field<object>("Revision") != null)
                                {
                                    row["revision"] = fr.Field<object>("Revision").ToString();
                                }
                            }
                        }
                        else
                        {
                            //so something with rows not found. 
                        }
                    }


                    MyConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileToUpdate + ";Extended Properties='Excel 12.0;HDR=YES';");
                    MyConnection.Open();
                    myCommand.Connection = MyConnection;
                    foreach (DataRow row in updateTable.Rows)
                    {
                        sql = $"Update ['{ILOX_PART_SHEET}$'] " +
                            $"Set [revision]='{row["revision"].ToString()}'" +
                            $" where " + (fileToUpdate.Contains("ToT_WTPart") ? "partNumber" : "ASSEMBLYPARTNUMBER") + $"='{row["partNumber"].ToString()}';";

                        myCommand.CommandText = sql;
                        myCommand.ExecuteNonQuery();
                    }
                    MyConnection.Close();
                }
                catch (Exception ex)
                {
                    UnProcessedFiles.Add(fileToUpdate);
                    throw;
                }
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            UpdateRowSource(fileToUpdate);
        }

        //RULE: if it's(ToT_WTPart partNumber) in the BOM(ToT_BOM) then Source = 1 otherwise = 2
        private void UpdateRowSource(string fileToUpdate)
        {
            string totBOMFile = @"C:\Users\mruiz\Documents\PLM\R2\Output\ToT_BOM\ToT_BOM_v1-020421.xlsx";

            //string copyFilePath = fileToUpdate;
            //int count = 0;
            //while (File.Exists(copyFilePath))
            //{
            //    count++;
            //    copyFilePath = Path.Combine(Path.GetDirectoryName(fileToUpdate), Path.GetFileNameWithoutExtension(fileToUpdate) + $"-Copy" + ((count > 0) ? count.ToString() : "") + "." + Path.GetExtension(fileToUpdate));
            //}
            //File.Copy(fileToUpdate, copyFilePath);

            ///////////////////////////////////////
            OleDbConnection oConn = null;
            OleDbCommand oComm = null;
            OleDbDataReader oRdr = null;
            string sConnString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileToUpdate + ";Extended Properties='Excel 12.0;HDR=YES;';";
            string sCommand = String.Empty;
            try
            {
                oConn = new OleDbConnection(sConnString);
                oConn.Open();

                //Set source as 1 for all partNumbers found in ToT_BOM and that are from China/Brazil
                sCommand = $"UPDATE ['{ILOX_PART_SHEET}$'] AS X " +
                           $"INNER JOIN (SELECT [ASSEMBLYPARTNUMBER] FROM [{BOM_SHEET}$] IN '{totBOMFile}' 'Excel 12.0;HDR=YES;') AS Y " +
                           $"ON X.[partNumber] = Y.[ASSEMBLYPARTNUMBER] " +
                           $"SET X.[source]= '1' " +
                           $"WHERE X.[FROM_SOURCE] IN ('A','B'); ";

                oComm = new OleDbCommand(sCommand, oConn);
                oRdr = oComm.ExecuteReader();

                //Set source as 2 for all partNumbers NOT found in ToT_BOM and are from China/Brazil
                sCommand = $"UPDATE ['{ILOX_PART_SHEET}$'] AS X " +
                           $"SET X.[source]= '2' " +
                           $"WHERE X.[FROM_SOURCE] IN ('A','B') AND X.[source] is null; ";

                oComm = new OleDbCommand(sCommand, oConn);
                oRdr = oComm.ExecuteReader();

                //set source and type to Make and Separable, respectively for any source equal to 1
                sCommand = $"UPDATE ['{ILOX_PART_SHEET}$'] AS X " +
                          $"SET X.[source]= 'Make', X.[type] = 'Separable' " +
                          $"WHERE X.[source] = '1'; ";

                oComm = new OleDbCommand(sCommand, oConn);
                oRdr = oComm.ExecuteReader();

                //set source and type to Buy and Component, respectively for all source values equal to 2
                sCommand = $"UPDATE ['{ILOX_PART_SHEET}$'] AS X " +
                          $"SET X.[source]= 'Buy', X.[type] = 'Component' " +
                          $"WHERE X.[source] = '2'; ";

                oComm = new OleDbCommand(sCommand, oConn);
                oRdr = oComm.ExecuteReader();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (oRdr != null) oRdr.Close();
                oRdr = null;
                oComm.Dispose();
                oConn.Close();
                oConn.Dispose();
            }


        }

        private void button25_Click(object sender, EventArgs e)
        {
            cadPartReportFile = GetFilePath();
            textBox12.Text = cadPartReportFile;

        }

        private void button26_Click(object sender, EventArgs e)
        {
            //Connection Strings
            string cadPartConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + cadPartReportFile + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1';";
            string totPartsConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + totPartsFile + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1';";
            string totEPMBuildRuleConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + totEPMBuildRuleFile + ";Extended Properties='Excel 12.0;HDR=YES;';";

            //SQL Strings
            string getCadNames = $"SELECT [CADNAME] FROM ['Export Worksheet$']";
            string getpartNamesBuy = $"SELECT [partNumber] FROM ['{ILOX_PART_SHEET}$'] WHERE source = 'Buy'";

            DataTable cadNameDT = GetDataTableFromExcel(cadPartConnectionString, getCadNames);
            DataTable partNameDT = GetDataTableFromExcel(totPartsConnectionString, getpartNamesBuy);

            //Return list of cad filename(with extension) and partName(filename without extenstion) from matching cad(without extension) partNumber in ToT_Parts(Buy only)
            var matchedCads = cadNameDT.AsEnumerable()
                .Join(partNameDT.AsEnumerable(),
                cad => Path.GetFileNameWithoutExtension(cad.Field<string>("CADNAME")),
                part => part.Field<string>("partNumber"),
                (cad, part) => new { cad = cad.Field<string>("CADNAME").ToString(), part = part.Field<string>("partNumber").ToString() }).ToList();

            string sql = string.Empty;
            int counter = 0;
            try
            {
                OleDbConnection MyConnection;
                OleDbCommand myCommand = new OleDbCommand();

                MyConnection = new OleDbConnection(totEPMBuildRuleConnectionString);
                MyConnection.Open();
                myCommand.Connection = MyConnection;

                //insert each row
                foreach (var row in matchedCads)
                {
                    sql = $"Insert into ['Export Worksheet$'] ([BUILDSOURCEOBJECTNUMBER],[BUILDTARGETOBJECTNUMBER]) values('{row.cad}','{row.part}');";
                    myCommand.CommandText = sql;
                    myCommand.ExecuteNonQuery();
                    counter++;
                }
                MyConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString() + "ROW: " + counter.ToString() + sql);
            }
        }

        private static DataTable GetDataTableFromExcel(string connectionString, string sqlString)
        {
            DataTable dt = new DataTable();

            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                try
                {
                    OleDbDataAdapter oleAdpt;
                    oleAdpt = new OleDbDataAdapter(sqlString, con); //here we read data from sheet1  
                    oleAdpt.Fill(dt); //fill excel data into dataTable  
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            return dt;
        }

        private DataTable ConvertToDataTable<TSource>(IEnumerable<TSource> source)
        {
            var props = typeof(TSource).GetProperties();

            var dt = new DataTable();
            dt.Columns.AddRange(
              props.Select(p => new DataColumn(p.Name, p.PropertyType)).ToArray()
            );

            source.ToList().ForEach(
              i => dt.Rows.Add(props.Select(p => p.GetValue(i, null)).ToArray())
            );

            return dt;
        }
        private void button27_Click(object sender, EventArgs e)
        {
            totPartsFile = GetFilePath();
            textBox13.Text = totPartsFile;
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            cadPartReportFile = ((TextBox)sender).Text;
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            totPartsFile = ((TextBox)sender).Text;
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            totEPMBuildRuleFile = ((TextBox)sender).Text;
        }

        private void button28_Click(object sender, EventArgs e)
        {
            totEPMBuildRuleFile = GetFilePath();
            textBox14.Text = totEPMBuildRuleFile;
        }
    }
}