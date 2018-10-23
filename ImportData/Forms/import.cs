using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;

namespace ImportData
{
    public partial class importDataForm : Form
    {
        public importDataForm()
        {
            InitializeComponent();
        }

        private void importDataForm_Load(object sender, EventArgs e)
        {

        }
        public string strFilePath;
        public DataTable GetCsv()
        {
            DataTable dt = new DataTable();
            using (StreamReader sr = new StreamReader(strFilePath))
            {
                string[] headers = sr.ReadLine().Split(',');
                foreach (string header in headers)
                {
                    dt.Columns.Add(header);
                }
                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(',');
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        dr[i] = rows[i];
                    }
                    dt.Rows.Add(dr);
                }
                return dt;

            }
        }

        private void txtFilePath_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtFilePath_MouseUp(object sender, MouseEventArgs e)
        {
        }
        private void GetDataFromExcel()
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            // dt.Column = colCount;  
            dataGridView.ColumnCount = colCount;
            dataGridView.RowCount = rowCount;

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {


                    //write the value to the Grid  


                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        dataGridView.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();
                    }
                    // Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");  

                    //add useful things here!     
                }
            }

            //cleanup  
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:  
            //  never use two dots, all COM objects must be referenced and released individually  
            //  ex: [somthing].[something].[something] is bad  

            //release com objects to fully kill excel process from running in the background  
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release  
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release  
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);


        }
       static string databasename = "";
        string fname = "";
        string databaseTable;
       // ArticleV2
 
        public IList<string> ListTables()
        {
            string myConnecting = @"Server= DESKTOP-5801JMQ\SQLEXPRESS01; Database= " + databasename + "; Integrated Security=True;";

            SqlConnection con = new SqlConnection(myConnecting);
            con.Open();
            List<string> tables = new List<string>();
            DataTable dt = con.GetSchema("Tables");
            foreach (DataRow row in dt.Rows)
            {
                string tablename = (string)row[2];
                tables.Add(tablename);
            }
            con.Close();
            return tables;
        }
        
        public DataTable getDataBaseData()
        {
            databaseTable = databaseTables.SelectedItem.ToString();


            var table = new DataTable();
            string myConnecting = @"Server= DESKTOP-5801JMQ\SQLEXPRESS01; Database= " + databasename + "; Integrated Security=True;";

            using (var da = new SqlDataAdapter("SELECT * FROM "+databaseTable, myConnecting))
            {
                da.Fill(table);
            }
            return table;

        }
        public void getDataBaseTables()
        {
            databaseTables.DataSource = ListTables();
        }
        private void button1_Click(object sender, EventArgs e)
            
        {
          
            if (txtType.SelectedItem.ToString() == "Excel File")
            {
                OpenFileDialog fdlg = new OpenFileDialog();
                fdlg.Title = "Excel File Dialog";
                fdlg.InitialDirectory = @"E:\";
                fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
                fdlg.FilterIndex = 2;
                fdlg.RestoreDirectory = true;
                if (fdlg.ShowDialog() == DialogResult.OK)
                {
                    fname = fdlg.FileName;
                    strFilePath = fdlg.FileName;

                }
                dataGridView.DataSource = null;

                GetDataFromExcel();

            }
            if (txtType.SelectedItem.ToString() == "CSV File")
            {
                OpenFileDialog fdlg = new OpenFileDialog();
                fdlg.Title = "Excel File Dialog";
                fdlg.InitialDirectory = @"c:\";
                fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
                fdlg.FilterIndex = 2;
                fdlg.RestoreDirectory = true;
                if (fdlg.ShowDialog() == DialogResult.OK)
                {
                    fname = fdlg.FileName;
                    strFilePath = fdlg.FileName;

                }
                dataGridView.DataSource = null;

                dataGridView.DataSource = GetCsv();
            }
            if (txtType.SelectedItem.ToString() == "DataBase File")
            {
                dataGridView.DataSource = null;

                dataGridView.DataSource = getDataBaseData();

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            databasename = txtdatabaseName.Text;
            getDataBaseTables();
            getDataBaseData();
        }
        private System.Windows.Forms.OpenFileDialog ofd;
        private System.Windows.Forms.FolderBrowserDialog fbd;
        private void button3_Click(object sender, EventArgs e)
        {

            try
            {
                //Build the CSV file data as a Comma separated string.
                string csv = string.Empty;

                //Add the Header row for CSV file.
                foreach (DataGridViewColumn column in dataGridView.Columns)
                {
                    csv += column.HeaderText + ',';
                }
                //Add new line.
                csv += "\r\n";

                //Adding the Rows

                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.Value != null)
                        {
                            //Add the Data rows.
                            csv += cell.Value.ToString().TrimEnd(',').Replace(",", ";") + ',';
                        }
                        // break;
                    }
                    //Add new line.
                    csv += "\r\n";
                }

                //Exporting to CSV.
                string folderPath = string.Empty;

                using (FolderBrowserDialog fdb = new FolderBrowserDialog())
                {
                    if (fdb.ShowDialog() == DialogResult.OK)
                    {
                        folderPath = fdb.SelectedPath + "\\";
                        MessageBox.Show(folderPath);
                    }
                }
        
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                File.WriteAllText(folderPath + "Exported.csv", csv);
                MessageBox.Show("done bb");
            }
            catch
            {
                MessageBox.Show("errror");
            }


        }

        private void copyAlltoClipboard()
        {
            dataGridView.SelectAll();
            DataObject dataObj = dataGridView.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            sfd.FileName = "Export.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                // Copy DataGridView results to clipboard
                copyAlltoClipboard();

                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlexcel = new Excel.Application();

                xlexcel.DisplayAlerts = false; // Without this you will get two confirm overwrite prompts
                Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                // Format column D as text before pasting results, this was required for my data
                Excel.Range rng = xlWorkSheet.get_Range("D:D").Cells;
                rng.NumberFormat = "@";

                // Paste clipboard results to worksheet range
                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                // For some reason column A is always blank in the worksheet. ¯\_(ツ)_/¯
                // Delete blank column A and select cell A1
                Excel.Range delRng = xlWorkSheet.get_Range("A:A").Cells;
                delRng.Delete(Type.Missing);
                xlWorkSheet.get_Range("A1").Select();

                // Save the excel file under the captured location from the SaveFileDialog
                xlWorkBook.SaveAs(sfd.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlexcel.DisplayAlerts = true;
                xlWorkBook.Close(true, misValue, misValue);
                xlexcel.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlexcel);

                // Clear Clipboard and DataGridView selection
                Clipboard.Clear();
                dataGridView.ClearSelection();

                // Open the newly saved excel file
                if (File.Exists(sfd.FileName))
                    System.Diagnostics.Process.Start(sfd.FileName);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string StrQuery;
            databasename = txtdatabaseName.Text;
            string ConnString = @"Server= DESKTOP-5801JMQ\SQLEXPRESS01; Database= " + databasename + "; Integrated Security=True;";

            try
            {
                using (SqlConnection conn = new SqlConnection(ConnString))
                {
                    using (SqlCommand comm = new SqlCommand())
                    {
                        comm.Connection = conn;
                        conn.Open();
                        for (int i = 0; i < dataGridView.Rows.Count; i++)
                        {
                            StrQuery = @"INSERT INTO "+ databaseTables.SelectedValue+ " VALUES ("
                                + dataGridView.Rows[i].Cells[1].Value + ", "
                                + dataGridView.Rows[i].Cells[2].Value + ", "
                                + dataGridView.Rows[i].Cells[3].Value + ", "
                                + dataGridView.Rows[i].Cells[4].Value + ", "
                                + dataGridView.Rows[i].Cells[5].Value + ", "
                                + dataGridView.Rows[i].Cells[6].Value + ");";
                            comm.CommandText = StrQuery;
                            comm.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            
        }
    }
}

