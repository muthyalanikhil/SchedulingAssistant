using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;

namespace SchedulingAssistant
{
    public partial class GenerateReport : Form
    {
        private DataGridView dataGridView;
        private List<int> infectedRowsList;
        private string reportBy;
        SchedulingAssistantHelper helper = new SchedulingAssistantHelper();

        public GenerateReport()
        {

        }

        public GenerateReport(DataGridView dataGridView, List<int> illegalRowsList)
        {
            InitializeComponent();
            reportOfCB.Enabled = false;
            GeneratePDFButton.Enabled = false;
            generateReportByCB.Text = "Please, select any value";
            generateReportByCB.Items.Add("Instructor");
            generateReportByCB.Items.Add("Room");
            generateReportByCB.Items.Add("All Instructors");
            this.dataGridView = dataGridView;
            this.infectedRowsList = illegalRowsList;
        }

        private void GeneratePDFButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (reportBy == "All Instructors")
                {
                    DataTable newDataTable = helper.MakeDataTable((DataTable)dataGridView.DataSource, generateReportByCB.Text, reportOfCB.Text, infectedRowsList);
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.InitialDirectory = "C";
                    saveFileDialog.Title = "Save an Excel File";
                    saveFileDialog.FileName = "Report";
                    saveFileDialog.Filter = "Excel Files(2013)|*xlsx";
                    if (saveFileDialog.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
                    {
                        using (XLWorkbook wb = new XLWorkbook())
                        {
                            var ws = wb.Worksheets.Add(newDataTable, "Report");
                            wb.SaveAs(saveFileDialog.FileName.ToString().Contains(".xlsx") ? saveFileDialog.FileName.ToString() : saveFileDialog.FileName.ToString() + ".xlsx");
                            MessageBox.Show("Report generated successfully.", "Generate Report", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else
                {
                    DataTable newDataTable = helper.MakeDataTable((DataTable)dataGridView.DataSource, generateReportByCB.Text, reportOfCB.Text, infectedRowsList);
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.InitialDirectory = "C";
                    saveFileDialog.Title = "Save an PDF File";
                    saveFileDialog.FileName = reportOfCB.Text.Replace(':', '-').Replace('(', ' ').Replace(')', ' ').Trim();
                    saveFileDialog.Filter = "Pdf Files|*.pdf";

                    if (saveFileDialog.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
                    {
                        String path = (saveFileDialog.FileName.ToString().Contains(".pdf") ? saveFileDialog.FileName.ToString() : saveFileDialog.FileName.ToString() + ".pdf");
                        helper.ExportDataTableToPdf(newDataTable, path, "Time Sheet of " + reportOfCB.Text);
                        MessageBox.Show("Report generated successfully.", "Generate Report", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }     
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message");
            }
        }

        private void generateReportByCB_SelectionChangeCommitted(object sender, EventArgs e)
        {

        }

        private void generateReportByCB_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                List<string> cbListItems = new List<string>();
                cbListItems.Clear();
                reportOfCB.Items.Clear();
                reportOfCB.Text = "Please, select any value";
                GeneratePDFButton.Enabled = false;
                DataTable dataTable = (DataTable)dataGridView.DataSource;
                reportOfCB.Enabled = true;
                reportBy = generateReportByCB.Text;
                if (generateReportByCB.Text == "Instructor")
                {
                    label2.Visible = true;
                    reportOfCB.Visible = true;
                    GeneratePDFButton.Text = "Generate PDF";
                    for (int currentRow = 0; currentRow < dataTable.Rows.Count; currentRow++)
                    {
                        if (!infectedRowsList.Contains(currentRow))
                        {
                            DataRow rowValue = dataTable.Rows[currentRow];
                            String name = rowValue["INSTRUCTOR NAME"].ToString();
                            String id = rowValue["INSTRUCTOR_ID"].ToString();
                            String value = name + "(ID: " + id + ")";
                            if (!cbListItems.Contains(value))
                            {
                                cbListItems.Add(value);
                            }
                        }
                    }
                    foreach (String value in cbListItems)
                    {
                        reportOfCB.Items.Add(value);
                    }
                }
                if (generateReportByCB.Text == "Room")
                {
                    label2.Visible = true;
                    reportOfCB.Visible = true;
                    GeneratePDFButton.Text = "Generate PDF";
                    for (int currentRow = 0; currentRow < dataTable.Rows.Count; currentRow++)
                    {
                        if (!infectedRowsList.Contains(currentRow))
                        {
                            DataRow rowValue = dataTable.Rows[currentRow];
                            String value = rowValue["LOCATION"].ToString();
                            if (!cbListItems.Contains(value))
                            {
                                cbListItems.Add(value);
                            }
                        }
                    }
                    foreach (String value in cbListItems)
                    {
                        reportOfCB.Items.Add(value);
                    }
                }
                if (generateReportByCB.Text == "All Instructors")
                {
                    GeneratePDFButton.Text = "Export to Excel";
                    label2.Visible = false;
                    reportOfCB.Visible = false;
                    GeneratePDFButton.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message");
            }
        }

        private void reportOfCB_SelectedIndexChanged(object sender, EventArgs e)
        {
            GeneratePDFButton.Enabled = true;
        }
    }
}
