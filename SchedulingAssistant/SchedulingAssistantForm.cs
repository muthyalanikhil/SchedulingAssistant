using System;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.OleDb;
using ClosedXML.Excel;
using System.Diagnostics;
using System.Drawing;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace SchedulingAssistant
{
    public partial class SchedulingAssistantForm : Form
    {
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string importFilePath = "";
        String connStringFinal = string.Empty;
        String selectedSheetName;
        List<Int32> rowMarkedAsDeleted = new List<int>();
        SchedulingAssistantHelper helper = new SchedulingAssistantHelper();
        List<int> illegalRowList = new List<int>();
        List<int> infectedRows = new List<int>();
        DataTable gridDataBackup = new DataTable();
        List<ChangedCellValue> changedCellList = new List<ChangedCellValue>();

        public SchedulingAssistantForm()
        {
            InitializeComponent();
            this.Location = new Point(0, 0);
            this.Size = Screen.PrimaryScreen.WorkingArea.Size;
            ImportFileButton.Enabled = false;
            CheckConflictButton.Enabled = false;
            GenerateReportButton.Enabled = false;
            AddRowButton.Enabled = false;
            deleteEmptyColumns.Enabled = false;
            deleteEmptyRows.Enabled = false;
            DeleteRowButton.Enabled = false;
        }

        private void SelectFileButton_Click(object sender, EventArgs e)
        {
            try
            {
                CloseAllExcelProcesses();
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog1.Title = "Select a Cursor File";

                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string filePath = openFileDialog1.FileName;
                    string extension = Path.GetExtension(filePath);
                    string header = "Yes";
                    string conStr;
                    conStr = string.Empty;
                    importFilePath = filePath;
                    switch (extension)
                    {
                        case ".xls": //Excel 97-03
                            conStr = string.Format(Excel03ConString, filePath, header);
                            connStringFinal = conStr;
                            break;

                        case ".xlsx": //Excel 07
                            conStr = string.Format(Excel07ConString, filePath, header);
                            connStringFinal = conStr;
                            break;
                    }

                    using (OleDbConnection con = new OleDbConnection(conStr))
                    {
                        using (OleDbCommand cmd = new OleDbCommand())
                        {
                            cmd.Connection = con;
                            con.Open();
                            DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            ExcelSheetNames.Items.Clear();
                            foreach (DataRow row in dtExcelSchema.Rows)
                            {
                                if (!row["TABLE_NAME"].ToString().Contains("FilterDatabase"))
                                {
                                    ExcelSheetNames.Items.Add(row["TABLE_NAME"].ToString().Trim().Replace("'", string.Empty).Replace("$", string.Empty));
                                }
                            }
                            ExcelSheetNames.SelectedIndex = 0;
                            excelFilePathTB.Text = openFileDialog1.FileName.ToString();
                            ExcelSheetNames.DropDownStyle = ComboBoxStyle.DropDownList;
                            con.Close();
                        }
                    }
                    ImportFileButton.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
            }
        }

        private void ImportFileButton_Click(object sender, EventArgs e)
        {
            try
            {
                rowMarkedAsDeleted.Clear();
                illegalRowList.Clear();
                infectedRows.Clear();
                gridDataBackup.Clear();
                changedCellList.Clear();

                CheckConflictButton.Enabled = true;
                GenerateReportButton.Enabled = true;
                AddRowButton.Enabled = true;
                deleteEmptyColumns.Enabled = true;
                deleteEmptyRows.Enabled = true;
                DeleteRowButton.Enabled = true;

                String connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + importFilePath + @";Extended Properties= ""Excel 12.0;HDR=YES;IMEX=1;MAXSCANROWS=15;READONLY=FALSE""";

                using (OleDbConnection con = new OleDbConnection(connectionString))
                {
                    using (OleDbCommand cmd = new OleDbCommand())
                    {
                        using (OleDbDataAdapter oda = new OleDbDataAdapter())
                        {
                            DataTable dt = new DataTable();

                            this.selectedSheetName = ExcelSheetNames.Text;
                            cmd.CommandText = "SELECT * From ['" + selectedSheetName + "$']";
                            cmd.Connection = con;
                            con.Open();
                            oda.SelectCommand = cmd;
                            dt.TableName = "test";
                            oda.Fill(dt);
                            con.Close();
                            List<String> columnNames = new List<string>();
                            columnNames.Add("DAYS");
                            columnNames.Add("TIME");
                            columnNames.Add("LOCATION");
                            columnNames.Add("INSTRUCTOR_ID");
                            columnNames.Add("SECTION");
                            columnNames.Add("CRSE#");
                            columnNames.Add("INSTRUCTOR NAME");
                            columnNames.Add("CRN");
                            bool isColumnValid = true;
                            String invalidColumn = "";
                            foreach (var item in columnNames)
                            {
                                if (!dt.Columns.Contains(item))
                                {
                                    isColumnValid = false;
                                    invalidColumn = invalidColumn + " " + item;
                                }
                            }

                            if (isColumnValid)
                            {
                                //Populate DataGridView.
                                dataGridView.DataSource = dt;
                                dataGridView.DoubleBuffered(true);
                                foreach (DataGridViewRow item in dataGridView.Rows)
                                {
                                    if (item.Index != -1)
                                    {
                                        foreach (DataGridViewCell cel in item.Cells)
                                        {
                                            string strSplitValue = Convert.ToString(cel.Value);
                                            if (strSplitValue.Contains("\n"))
                                            {
                                                changedCellList.Add(helper.updateDataBackup(cel.RowIndex, cel.ColumnIndex, strSplitValue.Split('\n')[1], strSplitValue.Split('\n')[0]));
                                                cel.Value = strSplitValue.Split('\n')[0];
                                            }
                                        }
                                    }
                                }
                                DataTable temporaryTable = dt;
                                for (int r = 0; r < dt.Rows.Count; r++)
                                {
                                    string value = dt.Rows[r]["CRN"].ToString();
                                    if (value.Contains("##"))
                                    {
                                        if (!rowMarkedAsDeleted.Contains(r))
                                        {
                                            rowMarkedAsDeleted.Add(r);
                                            value = value.Replace("##", string.Empty);
                                            dt.Rows[r]["CRN"] = value;
                                        }
                                    }
                                }
                                
                                //populate backup
                                gridDataBackup = dt;
                                foreach (var item in rowMarkedAsDeleted)
                                {
                                    dataGridView.Rows[item].DefaultCellStyle.Font = new System.Drawing.Font("Calibri", 11.23F, FontStyle.Strikeout);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Please check the imported excel file column names. Column name " + invalidColumn + " is/are in incorrect format or not found.", "Scheduling Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
            }
        }

        private void ExportFileButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (excelFilePathTB.Text == string.Empty || ExcelSheetNames.Text == String.Empty)
                {
                    MessageBox.Show("Please import an Excel file", "Scheduling Assistant", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (dataGridView.RowCount == 0)
                {
                    MessageBox.Show("Please import an Excel file", "Scheduling Assistant", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    Cursor.Current = Cursors.WaitCursor;
                    DataTable dt = (DataTable)(dataGridView.DataSource);

                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.InitialDirectory = "C";
                    saveFileDialog.Title = "Save an Excel File";
                    saveFileDialog.FileName = "MSACS";
                    saveFileDialog.Filter = "Excel Files(2013)|*xlsx";
                    if (saveFileDialog.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
                    {
                        using (XLWorkbook wb = new XLWorkbook())
                        {
                            var ws = wb.Worksheets.Add(gridDataBackup, ExcelSheetNames.Text + "Old");
                            wb.SaveAs(saveFileDialog.FileName.ToString().Contains(".xlsx") ? saveFileDialog.FileName.ToString() : saveFileDialog.FileName.ToString() + ".xlsx");
                        }
                        Excel.Application excelApp = new Excel.Application();

                        //Create an Excel workbook instance and open it from the predefined location
                        Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(saveFileDialog.FileName);

                        DataSet ds = new DataSet();
                        ds.Tables.Add(dt);
                        foreach (DataTable table in ds.Tables)
                        {
                            //Add a new worksheet to workbook with the Datatable name
                            Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                            excelWorkSheet.Name = ExcelSheetNames.Text;

                            for (int i = 1; i < table.Columns.Count + 1; i++)
                            {
                                excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                            }

                            for (int j = 0; j < table.Rows.Count; j++)
                            {
                                for (int k = 0; k < table.Columns.Count; k++)
                                {
                                    excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                                }
                            }
                            foreach (var item in changedCellList)
                            {
                                excelWorkSheet.Cells[item.row + 2, item.column + 1] = item.newValue + "\n" + item.oldValue;
                                Excel.Range range = excelWorkSheet.Cells[item.row + 2, item.column + 1];
                                if (item.newValue != null || item.oldValue != null)
                                {
                                    if (item.newValue == item.oldValue)
                                    {
                                        range.Value = item.newValue;
                                        range.get_Characters(1, item.newValue.Length).Font.Color = System.Drawing.Color.Black;
                                        range.get_Characters(1, item.newValue.Length).Font.Strikethrough = false;
                                    }
                                    else
                                    {
                                        range.get_Characters(1, item.newValue.Length + item.oldValue.Length + 1).Font.Color = System.Drawing.Color.Black;
                                        range.get_Characters(1, item.newValue.Length + item.oldValue.Length + 1).Font.Strikethrough = false;
                                        if (item.oldValue != null && item.oldValue.Length != 0)
                                        {
                                            range.get_Characters(1, item.newValue.Length).Font.Color = System.Drawing.Color.Red;
                                            range.get_Characters(item.newValue.Length + 1, item.oldValue.Length + 1).Font.Strikethrough = true;
                                        }
                                    }
                                }
                            }

                            foreach (var item in rowMarkedAsDeleted)
                            {
                                string value = dt.Rows[item]["CRN"].ToString();
                                excelWorkSheet.Cells[item + 2, 1] = "##" + value;
                                Excel.Range range = excelWorkSheet.Rows[item + 2];
                                range.Font.Strikethrough = true;
                            }
                        }
                        excelWorkBook.Save();
                        excelWorkBook.Close();
                        excelApp.Quit();
                        MessageBox.Show("File has been exported successfully.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    Cursor.Current = Cursors.Default;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
            }
        }

        private void CheckConflictButton_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = (DataTable)(dataGridView.DataSource);
                MarkIllegalRows(dt);

                DialogResult result = MessageBox.Show("Are you sure you want to continue conflict check without changes to below marked rows.", "Class Scheduling", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                   
                    List<int> finalRowsToIgnore = new List<int>();
                    if (rowMarkedAsDeleted.Count > 0)
                    {
                        foreach (var item in rowMarkedAsDeleted)
                        {
                            finalRowsToIgnore.Add(item);
                        }
                    }

                    if (illegalRowList.Count > 0)
                    {
                        foreach (var item in illegalRowList)
                        {
                            if (!finalRowsToIgnore.Contains(item))
                            {
                                finalRowsToIgnore.Add(item);
                            }
                        }
                        infectedRows = finalRowsToIgnore;
                    }
                                    
                    foreach (var row in helper.CheckConflict(dt, finalRowsToIgnore))
                    {
                        if (!infectedRows.Contains(row))
                        {
                            infectedRows.Add(row);
                        }
                        dataGridView.Rows[row].DefaultCellStyle.BackColor = Color.IndianRed;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
            }
        }

        private void MarkIllegalRows(DataTable dt)
        {
            try
            {
                if (illegalRowList.Count > 0)
                {
                    illegalRowList.Clear();
                }
                
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    dataGridView.Rows[r].DefaultCellStyle.BackColor = Color.White;
                }
                illegalRowList = helper.illegalFormatRowList(dt);
                foreach (var row in illegalRowList)
                {
                    dataGridView.Rows[row].DefaultCellStyle.BackColor = Color.LightBlue;
                }
                MessageBox.Show("Rows marked with blue are having cell data in wrong format. Please update them for including them in conflict check.", "Scheduling Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
            }
        }

        private void MarkEmptyCells(DataTable dt)
        {
            try
            {
                List<CellIndex> cellList = new List<CellIndex>();
                cellList = helper.emptyCellList(dt);
                foreach (var cell in cellList)
                {
                    dataGridView.Rows[cell.row].Cells[cell.column].Style.BackColor = Color.LightGray;
                }
                MessageBox.Show("Cells marked with gray are empty. Please update them for completing conflict check.", "Scheduling Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
            }
        }

        private void AddRowButton_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dataTable = (DataTable)dataGridView.DataSource;
                DataRow drToAdd = dataTable.NewRow();
                dataTable.Rows.InsertAt(drToAdd, dataGridView.RowCount);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
            }
        }

        private void DeleteRowButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView.SelectedCells.Count >= 1 && DeleteRowButton.Text == "Delete Row")
                {
                    DialogResult result = MessageBox.Show("Are you sure you want to delete record", "Class Scheduling", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        dataGridView.Rows[dataGridView.CurrentCell.RowIndex].DefaultCellStyle.Font = new System.Drawing.Font("Calibri", 11.23F, FontStyle.Strikeout);
                        rowMarkedAsDeleted.Add(dataGridView.CurrentCell.RowIndex);
                        DeleteRowButton.Text = "Undelete Row";
                    }
                }
                else
                {
                    dataGridView.Rows[dataGridView.CurrentCell.RowIndex].DefaultCellStyle.Font = new System.Drawing.Font("Calibri", 11.23F);
                    if (rowMarkedAsDeleted.Contains(dataGridView.CurrentCell.RowIndex))
                    {
                        rowMarkedAsDeleted.Remove(dataGridView.CurrentCell.RowIndex);
                        DeleteRowButton.Text = "Delete Row";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
            }
        }

        private void deleteEmptyRows_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = MessageBox.Show("Are you sure you want to delete all empty rows ?", "Scheduling Assistant", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    DataTable dataTable = (DataTable)dataGridView.DataSource;
                    bool flag = false;
                    for (int i = dataTable.Rows.Count - 1; i >= 0; i--)
                    {
                        DataRow dr = dataTable.Rows[i];
                        if (helper.AreAllColumnsEmpty(dr))
                        {
                            dr.Delete();
                            flag = true;
                        }
                    }
                    dataTable.AcceptChanges();
                    dataGridView.DataSource = dataTable;
                    if (flag)
                    {
                        MessageBox.Show("Rows have been deleted successfully.", "Scheduling Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Sorry. No empty rows to delete.", "Scheduling Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
            }
        }

        private void deleteEmptyColumns_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = MessageBox.Show("Are you sure you want to delete all empty columns ?", "Scheduling Assistant", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    DataTable dataTable = (DataTable)dataGridView.DataSource;
                    bool flag = helper.IsColumnEmpty(dataTable);
                    if (flag)
                    {
                        MessageBox.Show("Columns have been deleted successfully.", "Scheduling Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Sorry. No empty columns to delete.", "Scheduling Assistant", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
            }
        }

        private void GenerateReportButton_Click(object sender, EventArgs e)
        {
            try
            {
                GenerateReport generateReportForm = new GenerateReport(dataGridView, infectedRows);
                generateReportForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
            }
        }

        //Windows close button overridden. It kills all the Excel instances when closing the application
        protected override void OnFormClosing(System.Windows.Forms.FormClosingEventArgs e)
        {
            try
            {
                base.OnFormClosing(e);
                if (e.CloseReason == CloseReason.WindowsShutDown) return;
                Process[] allExclProc = Process.GetProcessesByName("EXCEL");
                foreach (Process allProcID in allExclProc)
                {
                    int procsID = allProcID.Id;
                    Process.GetProcessById(procsID).Kill();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
            }
            Clipboard.Clear();
            Application.Exit();
        }

        private void CloseAllExcelProcesses()
        {
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        string str = ex.StackTrace;
                        Console.WriteLine(str);
                    }
                }
            }
        }

        private void dataGridView_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                var oldValue = dataGridView[e.ColumnIndex, e.RowIndex].Value.ToString();
                var newCellValue = e.FormattedValue.ToString();
                if (oldValue != newCellValue)
                {
                    foreach (var item in changedCellList)
                    {
                        if (item.column == e.ColumnIndex && item.row == e.RowIndex)
                        {
                            changedCellList.Add(helper.updateDataBackup(e.RowIndex, e.ColumnIndex, item.oldValue, newCellValue));
                            return;
                        }
                    }
                    changedCellList.Add(helper.updateDataBackup(e.RowIndex, e.ColumnIndex, oldValue, newCellValue));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
            }
        }

        private void dataGridView_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (rowMarkedAsDeleted.Contains(dataGridView.CurrentCell.RowIndex))
            {
                DeleteRowButton.Text = "Undelete Row";
            }
            else
            {
                DeleteRowButton.Text = "Delete Row";
            }
        }

        private void dataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("Data entered is not valid for the cell. Please enter valid data.", "Data Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
