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
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Reflection.Emit;
using System.Reflection;


namespace ExcelApp
{
    public partial class Form4 : Form
    {
        private object lblArrow;
        public Excel._Worksheet xlWorksheet;
        public Excel.Workbook xlWorkBook;
        Microsoft.Office.Interop.Excel.Range xlRange;
        public Microsoft.Office.Interop.Excel.DataTable DataTable { get; }
        string connectionString = ConfigurationManager.ConnectionStrings["LMS_SQL"].ConnectionString;

        public string SingleFileName { get; private set; }

        string fileExcel;
        Excel.Application xlApp;
        private object label6;
        private object label7;
        private int intQNo;
        private string myName;
        private string myExcelPath;

        public Form4()
        {
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)

        {
            {
                try
                {
                    KillMe();
                    // SetHeight();
                    if (getLoginId() == false)
                    {
                        System.Environment.Exit(0);
                    }

                    string xx = GlobalVar.strSheetName;
                    int intQNos = intQNo;
                    //   button8.BackColor = Color.FromArgb(192, 192, 192);

                    string fullName = Assembly.GetEntryAssembly().Location;
                    myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
                    myExcelPath = System.IO.Path.GetDirectoryName(fullName);

                    string myDirName = System.IO.Directory.GetCurrentDirectory();
                    SingleFileName = "Financial_Services_student_book.xlsx";
                    fileExcel = myExcelPath + "\\Financial_Services_student_book.xlsx"; //"D:\\CharakPoints.xlsx";
                    FileInfo fileInfo = new FileInfo(fileExcel);
                    fileInfo.IsReadOnly = false;
                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(fileExcel, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    xlApp.DisplayAlerts = false;
                    xlApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;
                    xlApp.Visible = true;
                    getCaseStudy(Convert.ToInt32(myName));

                    if (frmUserLogin.strUserId != null)
                    {
                        label6 = frmUserLogin.strUserId;
                        label7 = frmUserLogin.strUserName;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Please close excel sheet first!");
                }
                System.Drawing.Rectangle screen = Screen.FromPoint(Cursor.Position).WorkingArea;

                // Calculate the width of the sidebar (20% of the screen width)
                int sidebarWidth = (int)(screen.Width * 0.23);

                // Set the sidebar width to 20% of the screen
                this.Width = sidebarWidth;

                // Set the sidebar height to the full screen height
                this.Height = screen.Height;

                // Position the sidebar on the right side of the screen
                this.Location = new System.Drawing.Point(screen.Width - sidebarWidth, screen.Top);

                // Optional: Set the StartPosition if you want to keep it consistent with your previous approach
                this.StartPosition = FormStartPosition.Manual;
                //kartik

            }
        }
        private void KillMe()
        {
            // Placeholder: Implement any cleanup or termination logic here
            // Example: Close the Excel application if it's running
            if (xlApp != null)
            {
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                xlApp = null;
            }
        }

        private bool getLoginId()
        {
            // Placeholder: Implement your login validation logic here
            // For now, let's assume login is successful
            return true;  // Or add logic to check login credentials
        }

        private void getCaseStudy(int v)
        {
            // Placeholder: Implement the logic for getting case studies based on the integer
            // Example: Retrieve case study data from an external source or database
          
        }

        private void RemoveReadOnlyAttribute(string fileExcel)
        {
            // This method removes the read-only attribute from a file if set
            FileInfo fileInfo = new FileInfo(fileExcel);
            if (fileInfo.IsReadOnly)
            {
                fileInfo.IsReadOnly = false;
            }
        }


        private void rjButton1_Click(object sender, EventArgs e)
        {
            Form2 frm = new Form2();

         
            frm.Show();

       
            this.Hide();
        }
        private void OpenExcel()
        {
            try
            {
                string fullName = Assembly.GetEntryAssembly().Location;
                string myExcelPath = Path.GetDirectoryName(fullName);
                string fileExcel = myExcelPath + "\\Financial_Services_student_book.xlsx";

                RemoveReadOnlyAttribute(fileExcel);

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(fileExcel, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlApp.DisplayAlerts = false;
                xlApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;
                xlApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening Excel file: " + ex.Message);
            }
        }




        //private void lblArrow_Click(object sender, EventArgs e)
        //{
        //    int screenWidth = Screen.PrimaryScreen.Bounds.Width;
        //    int newXPosition;

        //    // Initialize lblArrow to ">" or "<" somewhere before use
        //    if (lblArrow == null)
        //    {
        //        lblArrow = ">";
        //    }

        //    if (lblArrow.ToString() == ">")
        //    {
        //        newXPosition = screenWidth - 50;  // Move the form to the right
        //        lblArrow = "<";  // Change the direction

        //        foreach (Control control in this.Controls)
        //        {
        //            if (control != sender)  // Exclude the button (lblArrow)
        //            {
        //                control.Visible = true;
        //            }
        //        }
        //    }
        //    else
        //    {
        //        newXPosition = 1000;  // Move the form to the left (for example)
        //        lblArrow = ">";  // Change the direction

        //        foreach (Control control in this.Controls)
        //        {
        //            control.Visible = true;
        //        }
        //    }

        //    this.SetDesktopLocation(newXPosition, 0);  // Move form to the new X position
        //}

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void rjButton2_Click(object sender, EventArgs e)
        {
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int newXPosition;

            // Initialize lblArrow to ">" or "<" somewhere before use
            if (lblArrow == null)
            {
                lblArrow = ">";
            }

            if (lblArrow.ToString() == ">")
            {
                newXPosition = screenWidth - 50;  // Move the form to the right
                lblArrow = "<";  // Change the direction

                foreach (Control control in this.Controls)
                {
                    if (control != sender)  // Exclude the button (lblArrow)
                    {
                        control.Visible = true;
                    }
                }
            }
            else
            {
                newXPosition = 1000;  // Move the form to the left (for example)
                lblArrow = ">";  // Change the direction

                foreach (Control control in this.Controls)
                {
                    control.Visible = true;
                }
            }

            this.SetDesktopLocation(newXPosition, 0);  // Move form to the new X position
        }
    }
}
