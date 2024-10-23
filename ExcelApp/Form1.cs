using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
//using IronXL;
using System.Runtime.InteropServices;

namespace ExcelApp
{
    public partial class Form1 : Form
    {
        string fileExcel;
        Excel.Application xlApp;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {






            fileExcel = "D:\\TestExcel10.xlsx";

            //Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open("D:\\TestExcel1.xlsx");
            //Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

            //sheet.

           


            Excel.Workbook xlWorkBook;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fileExcel, 0,false , 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlApp.DisplayAlerts = false;
            xlApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;
            xlApp.Visible = true;
        }

        //private DataTable ReadExcel(string fileName)
        //{
        //    WorkBook workbook = WorkBook.Load(fileName);
        //    //// Work with a single WorkSheet.
        //    ////you can pass static sheet name like Sheet1 to get that sheet
        //    ////WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
        //    //You can also use workbook.DefaultWorkSheet to get default in case you want to get first sheet only
        //    WorkSheet sheet = workbook.DefaultWorkSheet;
        //    //Convert the worksheet to System.Data.DataTable
        //    //Boolean parameter sets the first row as column names of your table.
        //    return sheet.ToDataTable(true);
        //}

        private void button2_Click(object sender, EventArgs e)
        {
            //OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file
            //if (file.ShowDialog() == DialogResult.OK) //if there is a file chosen by the user
            //{
            ////string fileExt = Path.GetExtension("D:\\CharakPoints.xlsx"); //get the file extension
            ////if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
            ////{
            ////    try
            ////    {
            ////        DataTable dtExcel = ReadExcel("D:\\CharakPoints.xlsx"); //read excel file
            ////        dataGridView1.Visible = true;
            ////        dataGridView1.DataSource = dtExcel;
            ////    }
            ////    catch (Exception ex)
            ////    {
            ////        MessageBox.Show(ex.Message.ToString());
            ////    }
            ////}
            //else
            //{
            //    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
            //}
            //    }

            string fname = "D:\\CharakPoints.xlsx";

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            // dt.Column = colCount;  
            dataGridView1.ColumnCount = colCount;
            dataGridView1.RowCount = rowCount;

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {


                    //write the value to the Grid  


                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        dataGridView1.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();
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

        private void button4_Click(object sender, EventArgs e)
        {
            string fname = "D:\\Forecastingnew.xlsx";

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            // dt.Column = colCount;  
            dataGridView2.ColumnCount = colCount;
            dataGridView2.RowCount = rowCount;

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {


                    //write the value to the Grid  


                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        dataGridView2.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();
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
            //OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file
            //if (file.ShowDialog() == DialogResult.OK) //if there is a file chosen by the user
            //{
            ////string fileExt = Path.GetExtension("D:\\Forecastingnew.xlsx"); //get the file extension
            ////if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
            ////{
            ////    try
            ////    {
            ////        DataTable dtExcel = ReadExcel("D:\\Forecastingnew.xlsx"); //read excel file
            ////        dataGridView2.Visible = true;
            ////        dataGridView2.DataSource = dtExcel;
            ////    }
            ////    catch (Exception ex)
            ////    {
            ////        MessageBox.Show(ex.Message.ToString());
            ////    }
            ////}
            //    else
            //    {
            //        MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
            //    }
            //}
            //OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file
            //if (file.ShowDialog() == DialogResult.OK) //if there is a file chosen by the user
            //{
            //    string fileExt = Path.GetExtension(file.FileName); //get the file extension
            //    if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
            //    {
            //        try
            //        {
            //            DataTable dtExcel = ReadExcel(file.FileName); //read excel file
            //            dataGridView2.Visible = true;
            //            dataGridView2.DataSource = dtExcel;
            //        }
            //        catch (Exception ex)
            //        {
            //            MessageBox.Show(ex.Message.ToString());
            //        }
            //    }
            //    else
            //    {
            //        MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
            //    }
            //}
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form2 frmform2 = new Form2();
            frmform2.Close();
            int c = 2;
            int d = 2;
            dataGridView3.Rows.Clear();

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                //for (int j = 0; j < dataGridView2.RowCount-1; j++)
                //{
                string grid2 = dataGridView2.Rows[i].Cells[0].Value.ToString();
                string grid1 = dataGridView1.Rows[i].Cells[0].Value.ToString();

                var index = dataGridView3.Columns.Add("R", "Answer");
                var index2 = dataGridView3.Columns.Add("R1", "Status");
                var index1 = dataGridView3.Rows.Add();
                dataGridView3.Rows[index1].Cells["R"].Value = grid1;

                if (grid1 == grid2)
                {
                    Excel.Application xlApp = new Excel.Application();
                    fileExcel = "D:\\CharakPoints.xlsx";
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileExcel);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1]; // Insert your sheet index here
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    // Some code goes here
                    xlApp.DisplayAlerts = false;

                    // Update the excel worksheet
                    xlWorksheet.Cells[c, d] = "True";
                    xlWorksheet.SaveAs(fileExcel);
                    //xlWorksheet.Cells[1, 2] = "Value";
                    //xlWorksheet.SaveAs("TestExcel10.xlsx");
                    xlApp.Workbooks.Close();
                    xlApp.Quit();

                    Marshal.ReleaseComObject(xlWorksheet);
                    Marshal.ReleaseComObject(xlWorkbook);
                    Marshal.ReleaseComObject(xlApp);
                    dataGridView3.Rows[index1].Cells["R1"].Value = "True";

                }
                else
                {
                    Excel.Application xlApp = new Excel.Application();
                    fileExcel = "D:\\CharakPoints.xlsx";
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileExcel);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1]; // Insert your sheet index here
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    // Some code goes here


                    // Update the excel worksheet
                    xlApp.DisplayAlerts = false;
                    xlWorksheet.Cells[c, d] = "False";
                    //xlWorksheet.Cells[1, 2] = "Value";
                    xlWorksheet.SaveAs(fileExcel);
                    xlApp.Workbooks.Close();
                    xlApp.Quit();

                    Marshal.ReleaseComObject(xlWorksheet);
                    Marshal.ReleaseComObject(xlWorkbook);
                    Marshal.ReleaseComObject(xlApp);

                    dataGridView3.Rows[index1].Cells["R1"].Value = "False";
                }

                index++;
                index1++;
                index2++;
                c++;

            }

            //Form2 frmform2 = new Form2();
            //frmform2.Close();
            //int c = 2;
            //int d = 2;
            //dataGridView3.Rows.Clear();
            //for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            //{
            //    //for (int j = 0; j < dataGridView2.RowCount-1; j++)
            //    //{
            //    string grid2 = dataGridView2.Rows[i].Cells[0].Value.ToString();
            //    string grid1 = dataGridView1.Rows[i].Cells[0].Value.ToString();

            //    var index = dataGridView3.Columns.Add("R", "Answer");
            //    var index2 = dataGridView3.Columns.Add("R1", "Status");
            //    var index1 = dataGridView3.Rows.Add();
            //    dataGridView3.Rows[index1].Cells["R"].Value = grid1;

            //    if (grid1 == grid2)
            //    {
            //        Excel.Application xlApp = new Excel.Application();
            //        fileExcel = "D:\\CharakPoints.xlsx";
            //        Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileExcel);
            //        Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1]; // Insert your sheet index here
            //        Excel.Range xlRange = xlWorksheet.UsedRange;

            //        // Some code goes here
            //        xlApp.DisplayAlerts = false;

            //        // Update the excel worksheet
            //        xlWorksheet.Cells[c, d] = "True";
            //        xlWorksheet.SaveAs(fileExcel);
            //        //xlWorksheet.Cells[1, 2] = "Value";
            //        //xlWorksheet.SaveAs("TestExcel10.xlsx");
            //        xlApp.Workbooks.Close();
            //        xlApp.Quit();

            //        Marshal.ReleaseComObject(xlWorksheet);
            //        Marshal.ReleaseComObject(xlWorkbook);
            //        Marshal.ReleaseComObject(xlApp);
            //        dataGridView3.Rows[index1].Cells["R1"].Value = "True";

            //    }
            //    else
            //    {
            //        Excel.Application xlApp = new Excel.Application();
            //        fileExcel = "D:\\CharakPoints.xlsx";
            //        Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileExcel);
            //        Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1]; // Insert your sheet index here
            //        Excel.Range xlRange = xlWorksheet.UsedRange;

            //        // Some code goes here


            //        // Update the excel worksheet
            //        xlApp.DisplayAlerts = false;
            //        xlWorksheet.Cells[c, d] = "False";
            //        //xlWorksheet.Cells[1, 2] = "Value";
            //        xlWorksheet.SaveAs(fileExcel);
            //        xlApp.Workbooks.Close();
            //        xlApp.Quit();

            //        Marshal.ReleaseComObject(xlWorksheet);
            //        Marshal.ReleaseComObject(xlWorkbook);
            //        Marshal.ReleaseComObject(xlApp);

            //        dataGridView3.Rows[index1].Cells["R1"].Value = "False";
            //    }

            //    index++;
            //    index1++;
            //    index2++;
            //    c++;
            //    //xlApp.Workbooks.Close();
            //    //xlApp.Quit();
            //    //}

            //    //}
            //}
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            xlApp = null;
            getMethod();
            getMethod1();
        }

        private void getMethod()
        {
            //dataGridView1.Rows.Clear();
                string fileExt = Path.GetExtension("D:\\TestExcel10.xlsx"); //get the file extension
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    //try
                    //{
                    //    DataTable dtExcel = ReadExcel("D:\\TestExcel10.xlsx"); //read excel file
                    //    dataGridView1.Visible = true;
                    //    dataGridView1.DataSource = dtExcel;
                    //}
                    //catch (Exception ex)
                    //{
                    //    MessageBox.Show(ex.Message.ToString());
                    //}
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
                }
            
        }

        private void getMethod1()
        {
            //dataGridView2.Rows.Clear();
            string fileExt = Path.GetExtension("D:\\TestExcelTest.xlsx"); //get the file extension
            if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
            {
                //try
                //{
                //    DataTable dtExcel = ReadExcel("D:\\TestExcelTest.xlsx"); //read excel file
                //    dataGridView2.Visible = true;
                //    dataGridView2.DataSource = dtExcel;
                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show(ex.Message.ToString());
                //}
            }
            else
            {
                MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
           // OpenForm();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            
            Form2 login = new Form2();
            login.ShowDialog();
            Form1 login1 = new Form1();
            login1.Close();
            
        }

        //private void OpenForm(Form form)
        //{
        //    PositionReporterEdge(form);
        //    form.Show();
        //}

        ///// <summary>
        ///// Position the "Reporter" form next to the current form.
        ///// </summary>
        //private void PositionReporterEdge(Form form)
        //{
        //    int screenHeight = Screen.PrimaryScreen.WorkingArea.Height;
        //    int screenWidth = Screen.PrimaryScreen.WorkingArea.Width;

        //    Point parentPoint = this.Location;

        //    int parentHeight = this.Height;
        //    int parentWidth = this.Width;

        //    int childHeight = form.Height;
        //    int childWidth = form.Width;

        //    int resultX;
        //    int resultY;

        //    if ((parentPoint.Y + parentHeight + childHeight) > screenHeight)
        //    {
        //        // If we would move off the screen, position near the top.
        //        resultY = parentPoint.Y + 100; // move down 50
        //        resultX = parentPoint.X + 50;
        //    }
        //    else
        //    {
        //        // Position on the edge.
        //        resultY = parentPoint.Y + parentHeight;
        //        resultX = parentPoint.X;
        //    }

        //    // set our child form to the new position
        //    form.Location = new Point(resultX, resultY);
        //}

    }
}
