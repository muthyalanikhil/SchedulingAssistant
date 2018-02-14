using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;


namespace SchedulingAssistant
{
    public class SchedulingAssistantHelper
    {
        public List<int> CheckConflict(DataTable dt, List<int> illegalRowsList)
        {
            try
            {
                List<int> rowList = new List<int>();

                for (int currentRow = 0; currentRow < dt.Rows.Count; currentRow++)
                {
                    DataRow rowValue = dt.Rows[currentRow];
                    for (int otherRow = 0; otherRow < dt.Rows.Count; otherRow++)
                    {
                        DataRow otherRowValues = dt.Rows[otherRow];
                        if (rowValue != otherRowValues && !illegalRowsList.Contains(currentRow) && !illegalRowsList.Contains(otherRow))
                        {
                            String[] weekDay = new String[] { "M", "T", "W", "R", "F" };
                            foreach (var day in weekDay)
                            {
                                if (rowValue["DAYS"].ToString().Contains(day) && otherRowValues["DAYS"].ToString().Contains(day))
                                {
                                    if (IsTimeOverLapping(rowValue["TIME"].ToString(), otherRowValues["TIME"].ToString()))
                                    {
                                        if (rowValue["LOCATION"].ToString() == otherRowValues["LOCATION"].ToString() || rowValue["INSTRUCTOR_ID"].ToString() == otherRowValues["INSTRUCTOR_ID"].ToString() || (rowValue["SECTION"].ToString() + rowValue["CRSE#"].ToString()) == (otherRowValues["SECTION"].ToString() + otherRowValues["CRSE#"].ToString()))
                                        {
                                            if (!rowList.Contains(currentRow))
                                            {
                                                rowList.Add(currentRow);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                return rowList;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
                return null;
            }
        }

        public bool IsTimeOverLapping(string time1, string time2)
        {
            try
            {
                DateTime startTime1 = Convert.ToDateTime(time1.Split('-')[0].Insert(2, ":"));
                DateTime endtime1 = Convert.ToDateTime(time1.Split('-')[1].Insert(2, ":"));
                DateTime startTime2 = Convert.ToDateTime(time2.Split('-')[0].Insert(2, ":"));
                DateTime endtime2 = Convert.ToDateTime(time2.Split('-')[1].Insert(2, ":"));

                bool overlap = startTime1 <= endtime2 && startTime2 <= endtime1;
                return overlap;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
                return false;
            }
        }

        public bool AreAllColumnsEmpty(DataRow dr)
        {
            try
            {
                bool isEmpty = false;
                for (int i = 0; i < dr.ItemArray.Length; i++)
                {
                    if (dr.ItemArray[i].ToString() == "{}" || dr.ItemArray[i] == null || dr.ItemArray[i].ToString() == string.Empty)
                    {
                        isEmpty = true;
                    }
                    else
                    {
                        isEmpty = false;
                        return isEmpty;
                    }
                }
                return isEmpty;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
                return false;
            }
        }

        public bool IsColumnEmpty(DataTable dt)
        {
            try
            {
                bool isEmpty = false;
                bool flag = false;
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        if (dt.Rows[row][col].ToString() == "{}" || dt.Rows[row][col] == null || dt.Rows[row][col].ToString() == string.Empty)
                        {
                            isEmpty = true;
                        }
                        else
                        {
                            isEmpty = false;
                            break;
                        }
                    }
                    if (isEmpty == true)
                    {
                        dt.Columns.Remove(dt.Columns[col]);
                        flag = true;
                    }
                }
                return flag;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
                return false;
            }
        }

        public List<CellIndex> emptyCellList(DataTable dt)
        {
            try
            {
                List<CellIndex> index = new List<CellIndex>();

                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        if (isCellEmpty(dt, r, c))
                        {
                            index.Add(new CellIndex() { row = r, column = c });
                        }
                    }

                }
                return index;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
                return null;
            }
        }

        public bool isCellEmpty(DataTable dt, int row, int column)
        {
            try
            {
                bool isCellEmpty = false;
                if (dt.Rows[row][column].ToString() == "{}" || dt.Rows[row][column] == null || dt.Rows[row][column].ToString() == string.Empty)
                {
                    isCellEmpty = true;
                    return isCellEmpty;
                }
                return isCellEmpty;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
                return false;
            }
        }

        public bool matchString(String reg, String text)
        {
            try
            {
                Regex r = new Regex(reg);
                Match m = r.Match(text);
                if (m.Success)
                {
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
                return false;
            }
        }

        public List<int> illegalFormatRowList(DataTable dt)
        {
            try
            {
                List<int> rowList = new List<int>();
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    if (!matchString(@"^\d{4}-\d{4}$", dt.Rows[r]["TIME"].ToString()) || !matchString(@"M|T|W|R|F|MT|MW|MR|MF|TW|TR|TF|WR|WF|RF|MTW|MTR|MTF|MWR|MWF|MRF|TWR|TWF|TRF|WRF", dt.Rows[r]["DAYS"].ToString()))
                    {
                        if (!rowList.Contains(r))
                        {
                            rowList.Add(r);
                        }
                    }
                    DataRow rowValue = dt.Rows[r];
                    if (isCellEmpty(dt, r, rowValue.Table.Columns["INSTRUCTOR_ID"].Ordinal))
                    {
                        if (!rowList.Contains(r))
                        {
                            rowList.Add(r);
                        }
                    }
                    if (isCellEmpty(dt, r, rowValue.Table.Columns["LOCATION"].Ordinal))
                    {
                        if (!rowList.Contains(r))
                        {
                            rowList.Add(r);
                        }
                    }
                    if (isCellEmpty(dt, r, rowValue.Table.Columns["INSTRUCTOR NAME"].Ordinal))
                    {
                        if (!rowList.Contains(r))
                        {
                            rowList.Add(r);
                        }
                    }
                }
                return rowList;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
                return null;
            }
        }

        public void ExportDataTableToPdf(DataTable dtblTable, String strPdfPath, string strHeader)
        {
            try
            {
                System.IO.FileStream fs = new FileStream(strPdfPath, FileMode.Create, FileAccess.Write, FileShare.None);
                Document document = new Document();
                document.SetPageSize(iTextSharp.text.PageSize.A4);
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                document.Open();

                //Report Header
                BaseFont bfntHead = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                Font fntHead = new Font(bfntHead, 16, 1, BaseColor.GRAY);
                Paragraph prgHeading = new Paragraph();
                prgHeading.Alignment = Element.ALIGN_CENTER;
                prgHeading.Add(new Chunk(strHeader.ToUpper(), fntHead));
                document.Add(prgHeading);

                //Author
                Paragraph prgAuthor = new Paragraph();
                BaseFont btnAuthor = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                Font fntAuthor = new Font(btnAuthor, 8, 2, BaseColor.GRAY);
                prgAuthor.Alignment = Element.ALIGN_RIGHT;
                prgAuthor.Add(new Chunk("Author : Scheduling Assistant", fntAuthor));
                prgAuthor.Add(new Chunk("\nRun Date : " + DateTime.Now.ToShortDateString(), fntAuthor));
                document.Add(prgAuthor);

                //Add a line seperation
                Paragraph p = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, BaseColor.BLACK, Element.ALIGN_LEFT, 1)));
                document.Add(p);

                //Add line break
                document.Add(new Chunk("\n", fntHead));

                //Write the table
                PdfPTable table = new PdfPTable(dtblTable.Columns.Count);
                //Table header
                BaseFont btnColumnHeader = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                Font fntColumnHeader = new Font(btnColumnHeader, 10, 1, BaseColor.WHITE);
                for (int i = 0; i < dtblTable.Columns.Count; i++)
                {
                    PdfPCell cell = new PdfPCell();
                    cell.BackgroundColor = BaseColor.GRAY;
                    cell.AddElement(new Chunk(dtblTable.Columns[i].ColumnName.ToUpper(), fntColumnHeader));
                    table.AddCell(cell);
                }
                //table Data
                for (int i = 0; i < dtblTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dtblTable.Columns.Count; j++)
                    {
                        table.AddCell(dtblTable.Rows[i][j].ToString());
                    }
                }
                document.Add(table);
                document.Close();
                writer.Close();
                fs.Close();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public DataTable MakeDataTable(DataTable dt, String type, String value, List<int> infectedRowsList)
        {
            try
            {
                DataTable newDT = new DataTable();
                switch (type)
                {
                    case "Room":
                        newDT.Columns.Add("Day");
                        newDT.Columns.Add("TIME");
                        newDT.Columns.Add("CRSE#");
                        newDT.Columns.Add("SECTION");
                        newDT.Columns.Add("INSTRUCTOR");
                        newDT.Columns.Add("TITLE");
                        for (int currentRow = 0; currentRow < dt.Rows.Count; currentRow++)
                        {
                            DataRow rowValue = dt.Rows[currentRow];
                            String excelValue = rowValue["LOCATION"].ToString();
                            if (excelValue == value && !infectedRowsList.Contains(currentRow))
                            {
                                newDT.Rows.Add(rowValue["DAYS"].ToString(), rowValue["TIME"].ToString(), rowValue["CRSE#"].ToString(), rowValue["SECTION"].ToString(), rowValue["INSTRUCTOR NAME"].ToString(), rowValue["TITLE"].ToString());
                            }
                        }
                        break;
                    case "Instructor":
                        newDT.Columns.Add("Day");
                        newDT.Columns.Add("TIME");
                        newDT.Columns.Add("CRSE#");
                        newDT.Columns.Add("LOCATION");
                        newDT.Columns.Add("TITLE");
                        for (int currentRow = 0; currentRow < dt.Rows.Count; currentRow++)
                        {
                            DataRow rowValue = dt.Rows[currentRow];
                            String name = rowValue["INSTRUCTOR NAME"].ToString();
                            String id = rowValue["INSTRUCTOR_ID"].ToString();
                            String excelValue = name + "(ID: " + id + ")";
                            if (excelValue == value && !infectedRowsList.Contains(currentRow))
                            {
                                newDT.Rows.Add(rowValue["DAYS"].ToString(), rowValue["TIME"].ToString(), rowValue["CRSE#"].ToString(), rowValue["LOCATION"].ToString(), rowValue["TITLE"].ToString());
                            }
                        }
                        break;
                    case "All Instructors":
                        infectedRowsList.Sort();
                        infectedRowsList.Reverse();
                        foreach (var item in infectedRowsList)
                        {
                            if (dt.Rows.Count > item)
                            {
                                dt.Rows.RemoveAt(item);
                            }                        
                        }
                        List<string> instructorsList = new List<string>(dt.Rows.Count);
                        foreach (DataRow row in dt.Rows)
                        {
                            if (!instructorsList.Contains((string)row["INSTRUCTOR NAME"]))
                            {
                                instructorsList.Add((string)row["INSTRUCTOR NAME"]);
                            }
                        }
                        dt.Columns["INSTRUCTOR NAME"].ColumnName = "INSTRUCTOR_NAME";
                        List<AllInstructorSections> allInstructorSectionList = new List<AllInstructorSections>();
                        allInstructorSectionList.Clear();
                        int numberOfColumns = 0;
                        foreach (var item in instructorsList)
                        {
                            AllInstructorSections data = new AllInstructorSections();
                            data.instructorName = item;
                            var list = GetRowsByFilter(dt, item);
                            list.Insert(0, item);
                            data.sectionList = list;
                            allInstructorSectionList.Add(data);
                            if (numberOfColumns < list.Count)
                            {
                                numberOfColumns = list.Count;
                            }
                        }
                        newDT.Columns.Add("Instructor");
                        for (int i = 1; i < numberOfColumns; i++)
                        {
                            newDT.Columns.Add("CrseSec-0" + i);
                        }
                        for (int i = 0; i < allInstructorSectionList.Count; i++)
                        {
                            newDT.Rows.Add(allInstructorSectionList[i].sectionList.ToArray());
                        }
                        dt.Columns["INSTRUCTOR_NAME"].ColumnName = "INSTRUCTOR NAME";
                        break;
                    default:
                        newDT.Columns.Add("No Data to display");
                        newDT.Rows.Add("No data to display");
                        break;
                }
                return newDT;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
                return null;
            }
        }

        private List<string> GetRowsByFilter(DataTable dt, String name)
        {
            // Presuming the DataTable has a column named Date.
            string expression;
            expression = "INSTRUCTOR_NAME = '" + name + "'";
            DataRow[] foundRows;

            DataView dv = new DataView(dt);
            dv.RowFilter = expression;

            // Use the Select method to find all rows matching the filter.
            foundRows = dt.Select(expression);
            List<string> crseList = new List<string>();
            // Print column 0 of each returned row.
            int crseColumnIndex = dt.Columns.IndexOf("CRSE#");
            int sectionColumnIndex = dt.Columns.IndexOf("SECTION");
            for (int i = 0; i < foundRows.Length; i++)
            {
                crseList.Add(foundRows[i][crseColumnIndex].ToString() + "-" + foundRows[i][sectionColumnIndex].ToString());
            }
            return crseList;
        }

        public ChangedCellValue updateDataBackup(int row, int column, String oldValue, String newCellValue)
        {
            try
            {
                ChangedCellValue updatedValues = new ChangedCellValue();
                updatedValues.row = row;
                updatedValues.column = column;
                updatedValues.oldValue = oldValue;
                updatedValues.newValue = newCellValue;
                return updatedValues;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                string str = ex.StackTrace;
                Console.WriteLine(str);
                return null;
            }
        }

        public DataTable ReportToExcel(DataTable dt)
        {

            return dt;
        }
    }
    public class CellIndex
    {
        public int row { get; set; }
        public int column { get; set; }
    }

    public class ChangedCellValue
    {
        public int row { get; set; }
        public int column { get; set; }
        public String oldValue { get; set; }
        public String newValue { get; set; }
    }

    public class AllInstructorSections
    {
        public String instructorName { get; set; }
        public List<string> sectionList { get; set; }
    }
}
