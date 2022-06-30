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
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using iTextSharp.text.pdf;
using iTextSharp.text;

namespace Fifa
{
    public partial class Form1 : Form
    {
        SqlConnection conn;
        string connString = @"Data Source=DESKTOP-UAK5LC3\SQLEXPRESS;Initial Catalog=Fifa;Integrated Security=True";
        //conn = new SqlConnection(connString);
        SqlCommand cmd;
        SqlDataAdapter adpt;
        DataTable dt;
        int Player_Id;

        public Form1()
        {
            InitializeComponent();
            displayInfo();
        }

        public static void ExportExcel(string fileName, DataGridView dataGridView1)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                string saveFileName = "";
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.DefaultExt = "xls";
                saveDialog.Filter = "Excel File|*.xls";
                saveDialog.FileName = fileName;
                saveDialog.ShowDialog();
                saveFileName = saveDialog.FileName;
                if (saveFileName.IndexOf(":") < 0) return;
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
                Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];

                //header
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    worksheet.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
                }
                //storing each row and column value for excel sheet
                for (int r = 0; r < dataGridView1.Rows.Count; r++)
                {
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        worksheet.Cells[r + 2, i + 1] = dataGridView1.Rows[r].Cells[i].Value;
                    }
                    Application.DoEvents();
                }
                worksheet.Columns.EntireColumn.AutoFit();

                if (saveFileName != "")
                {
                    try
                    {
                        workbook.Saved = true;
                        workbook.SaveCopyAs(saveFileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("The file may be opened now, Data not export\n" + ex.Message);
                    }

                }
                xlApp.Quit();
                MessageBox.Show("File save successful", "prompt", MessageBoxButtons.OK);
            }
            else
            {
                MessageBox.Show("the datagridview is empty", "prompt", MessageBoxButtons.OK);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            //SqlConnection conn;
            //string connString = @"Data Source=DESKTOP-UAK5LC3\SQLEXPRESS;Initial Catalog=Fifa;Integrated Security=True";
            conn = new SqlConnection(connString);

            conn.Open();
            cmd = new SqlCommand("insert into player1 values('" + txtPlayerName.Text + "','" + txtPlayerClub.Text + "','" + txtCountry.Text + "','" + txtJersyNo.Text + "')", conn);
            cmd.ExecuteNonQuery();
            MessageBox.Show("Player Data Stored in the DataBase");
            //displayInfo();
            conn.Close();
            displayInfo();
            Clear();

            /*
            System.Data.OleDb.OleDbConnection MyConnection;
            System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();
            string sql = null;
            MyConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\\Users\\sangr\\OneDrive\\Desktop\\Fifa\\Excel\\Fifa.xls';Extended Properties=Excel 8.0;");
            MyConnection.Open();
            myCommand.Connection = MyConnection;
            sql = "Insert into [Sheet1$] (id,name)  values('5','e')";
            myCommand.CommandText = sql;
            myCommand.ExecuteNonQuery();
            MyConnection.Close();*/
        }

        public void displayInfo()
        {
            conn = new SqlConnection(connString);
            conn.Open();
            adpt = new SqlDataAdapter("select * from player1", conn);
            dt = new DataTable();
            adpt.Fill(dt);
            dataGridView1.DataSource = dt;
            conn.Close();

        }

        public void Clear()
        {
            txtPlayerName.Text = "";
            txtPlayerClub.Text = "";
            txtCountry.Text = "";
            txtJersyNo.Text = "";
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Player_Id = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
            txtPlayerName.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            txtPlayerClub.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            txtCountry.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            txtJersyNo.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();

        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {

            conn = new SqlConnection(connString);
            conn.Open();
            cmd = new SqlCommand("update player1 set P_Name ='" + txtPlayerName.Text + "',Club='" + txtPlayerClub.Text + "',Country='" + txtCountry.Text + "',J_No='" + txtJersyNo.Text + "' where Player_Id='" + Player_Id + "'", conn);
            cmd.ExecuteNonQuery();
            MessageBox.Show("Player Information Has Been Updated !");
            //displayInfo();
            Clear();
            conn.Close();
            displayInfo();

        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            conn = new SqlConnection(connString);
            conn.Open();
            cmd = new SqlCommand("delete from player1 where Player_Id = '" + Player_Id + "'", conn);
            cmd.ExecuteNonQuery();
            MessageBox.Show("Player Data Deleted");
            conn.Close();
            displayInfo();
        }

        private void txtUserName_TextChanged(object sender, EventArgs e)
        {

        }

        private void btExcel_Click(object sender, EventArgs e)
        {
            string str = "Data";
            ExportExcel(str, dataGridView1);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            string path = @"C:\Users\sangr\OneDrive\Desktop\Fifa\Excel\Fifa.xls";
            xlWorkBook.SaveAs(path, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            MessageBox.Show("Excel file created");*/
        }

        private void btPdf_Click(object sender, EventArgs e)
        {
            /*PdfPTable table = new PdfPTable(dataGridView1.Columns.Count);

            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                table.AddCell(new Phrase(dataGridView1.Columns[j].HeaderText));
            }

            table.HeaderRows = 1;

            for(int i = 0 ; i < dataGridView1.Rows.Count; i++)
            {
                for(int k = 0; k < dataGridView1.Columns.Count; k++)
                {
                    if(dataGridView1[k,i].Value!=null)
                    {
                        table.AddCell(new Phrase(dataGridView1[k, i].Value.ToString()));
                    }
                }
            }*/
        }

        private void btCsv_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.InitialDirectory = "C";
            dlg.Title = "SaveAs";
            dlg.FileName = "";
            dlg.Filter = "CSV (Comma delimited)|*.csv";
            if (dlg.ShowDialog() != DialogResult.Cancel)
            {
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelApp.Application.Workbooks.Add(Type.Missing); //creating workbook

                ExcelApp.Columns.ColumnWidth = 20;
                //header
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    ExcelApp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }
                //storing each row and column value for excel sheet
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value != null)
                        {
                            ExcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }

                    }
                }
                if (dlg.FileName != "")
                {
                    try
                    {
                        ExcelApp.ActiveWorkbook.SaveCopyAs(dlg.FileName.ToString());
                        ExcelApp.ActiveWorkbook.Saved = true;
                        ExcelApp.Quit();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("export error, the file may be opened now\n" + ex.Message);
                    }

                }
                MessageBox.Show("File save successful", "prompt", MessageBoxButtons.OK);
            }
        }
    }
}


