using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Net;
using Application = System.Windows.Forms.Application;
using System.IO;
using System.Xml.Linq;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices.ComTypes;
using System.Diagnostics;
using Microsoft.Vbe.Interop;
using System.Xml;
using System.Reflection.Emit;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using static System.Net.WebRequestMethods;
using System.Security.Cryptography;
using Microsoft.Office.Core;
using static Class1;
using SixLabors.ImageSharp.Drawing;
using System.Drawing.Drawing2D;

namespace ExcelApp
{

    public partial class Form2 : Form
    {
        TimeSpan DtStart;
        TimeSpan DtFinal;
        int QNo = 0;
        int intCaseStudyId = 0;

        byte[] _barray1;
        byte[] _barray2;
        int currentRow = 0;
        int currentRow1 = 0;
        public Excel._Worksheet xlWorksheet;
        public Excel.Workbook xlWorkBook;
        Microsoft.Office.Interop.Excel.Range xlRange;
        public static string Qhint;
        public static string QLink;
        public static string strCell1;
        public static string strCell2;
        public string FilePath;
        //  public static string strSheetName;
        public string strUserName;
        public string strEmailId;
        public string SingleFileName;
        public int intUserId;
        public int SingleUserId;
        string myExcelPath;

        public string AnsStatus = "";
        public static CompareResult xt;
        string myName = "";
        int intQNo;
        SqlConnection conn;
        FileStream fstream;
        System.Data.DataTable dt;
        public Microsoft.Office.Interop.Excel.DataTable DataTable { get; }

        int val;
        int a;
        public int xx = 1;

        string connectionString = ConfigurationManager.ConnectionStrings["LMS_SQL"].ConnectionString;
        string fileExcel;
        Excel.Application xlApp;
        public Form2()
        {
            InitializeComponent();

        }
        public enum CompareResult
        {
            ciCompareOk,
            ciPixelMismatch,
            ciSizeMismatch
        };
        public static CompareResult Compare(Bitmap bmp1, Bitmap bmp2)
        {
            CompareResult cr = CompareResult.ciCompareOk;
            if (bmp1.Size != bmp2.Size)
            {
                cr = CompareResult.ciSizeMismatch;
            }
            else
            {
                System.Drawing.ImageConverter ic = new System.Drawing.ImageConverter();
                byte[] btImage1 = new byte[1];
                btImage1 = (byte[])ic.ConvertTo(bmp1, btImage1.GetType());
                byte[] btImage2 = new byte[1];
                btImage2 = (byte[])ic.ConvertTo(bmp2, btImage2.GetType());

                SHA256Managed shaM = new SHA256Managed();
                byte[] hash1 = shaM.ComputeHash(btImage1);
                byte[] hash2 = shaM.ComputeHash(btImage2);

                for (int i = 0; i < hash1.Length && i < hash2.Length && cr == CompareResult.ciCompareOk; i++)
                {
                    if (hash1[i] != hash2[i])
                        cr = CompareResult.ciPixelMismatch;
                }
                shaM.Clear();
            }

            return cr;
        }
        protected void SetHeight()
        {
            StartPosition = FormStartPosition.Manual;
            System.Drawing.Rectangle screen = Screen.FromPoint(Cursor.Position).WorkingArea;
            int w = Width >= screen.Width ? screen.Width : (screen.Width + Width) / 2;
            int h = Height >= screen.Height ? screen.Height : (screen.Height + Height) / 2;
            Location = new System.Drawing.Point(this.Width + 150, screen.Top + (screen.Height - h) / 2);
            this.Height = screen.Height;
        }
        protected void ResetHeight()
        {
            StartPosition = FormStartPosition.Manual;
            System.Drawing.Rectangle screen = Screen.FromPoint(Cursor.Position).WorkingArea;
            int w = Width >= screen.Width ? screen.Width : (screen.Width + Width) / 2;
            int h = Height >= screen.Height ? screen.Height : (screen.Height + Height) / 2;
            Location = new System.Drawing.Point(this.Width + 150, screen.Top + (screen.Height - h) / 2);
            this.Height = screen.Height;
        }
        protected void KillMe()
        {
            foreach (var process in Process.GetProcessesByName("Microsoft Excel (32 bit)"))
            {
                process.Kill();
            }
        }
        private void Form2_Load(object sender, EventArgs e)
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
                    label6.Text = frmUserLogin.strUserId;
                    label7.Text = frmUserLogin.strUserName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please close excel sheet first!");
            }
            System.Drawing.Rectangle screen = Screen.FromPoint(Cursor.Position).WorkingArea;

            // Calculate the width of the sidebar (20% of the screen width)
            int sidebarWidth = (int)(screen.Width * 0.2);

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
        private void RemoveReadOnlyAttribute(string filePath)
        {
            FileAttributes attributes = System.IO.File.GetAttributes(filePath);
            {
                System.IO.File.SetAttributes(filePath, attributes & ~FileAttributes.ReadOnly);
            }

        }
        private void OpenExcel()
        {
            string fullName = Assembly.GetEntryAssembly().Location;
            string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
            string myExcelPath = System.IO.Path.GetDirectoryName(fullName);
            fileExcel = myExcelPath + "\\Financial_Services_student_book.xlsx";

            RemoveReadOnlyAttribute(fileExcel);

            Excel.Workbook xlWorkBook;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fileExcel, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlApp.DisplayAlerts = false;
            xlApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;
            xlApp.Visible = true;

        }
        private bool getLoginId()
        {

            try
            {
                string fullName = Assembly.GetEntryAssembly().Location;
                //string myName = "1";

                strEmailId = GlobalVar.GEmailId;
                //myName =System.IO.Path.GetFileNameWithoutExtension(fullName);
                //intUserId = Convert.ToInt32(myName);
                //SqlConnection conn = new SqlConnection(connectionString);
                //conn.Open();
                //string selectquery = "SELECT * FROM UserMaster where PK_UserId = " + myName.ToString() + "";
                //SqlCommand cmd = new SqlCommand(selectquery, conn);
                //SqlDataReader reader1;
                //reader1 = cmd.ExecuteReader();
                //if (reader1.Read())
                //{
                //    strEmailId = reader1.GetValue(5).ToString();
                //    strUserName = reader1.GetValue(2).ToString();

                return true;

                //}
                //else
                //{
                //    MessageBox.Show("Invalid User!");
                //    return false;
                //}

            }
            catch (Exception ex)
            {

            }
            return false;

        }
        private void getStatus()
        {
            string str3 = "Select * from AnswerDetail where UserName = '" + label7.Text + "'";
            SqlDataAdapter sda = new SqlDataAdapter(str3, connectionString);
            dt = new System.Data.DataTable();
            sda.Fill(dt);

            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {

                    if (currentRow < dt.Rows.Count)
                    {
                        var row = dt.Rows[currentRow];
                        label1.Text = row["Question"].ToString();
                        label2.Text = "Question" + " " + (currentRow1 + 1);
                        currentRow1++;
                        incrementPbar();

                    }
                }
            }


        }

        private void getCaseStudy(int UserId)
        {
            int CaseStudyId = 0;

            ddlTestSheet.Items.Clear();
            conn = new SqlConnection();
            conn.ConnectionString = connectionString;
            strEmailId = GlobalVar.GEmailId;

            SqlDataAdapter das = new SqlDataAdapter("select top 1* from intallium_sa.Login where EmailId='" + strEmailId + "' and CaseStudyId=" + UserId + " order by sno desc", conn);
            DataSet dss = new DataSet();
            das.Fill(dss);
            if (dss.Tables[0].Rows.Count > 0)
            {

                intCaseStudyId = Convert.ToInt32(dss.Tables[0].Rows[0]["CaseStudyId"].ToString() + "");

                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
                conn.Open();

                SqlCommand cmd = new SqlCommand();

                cmd = conn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                string xName = "";
                cmd.CommandText = "select top 1 SheetName,QNo,Question from QuestionDetail where FK_CaseStudyId=" + intCaseStudyId + " order by QNo";
                cmd.ExecuteNonQuery();
                dt = new System.Data.DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    ddlTestSheet.Items.Add(dr["SheetName"].ToString());
                    // ddlTestSheet.Text = dr["SheetName"].ToString()+"";
                    xName = dr["SheetName"].ToString() + "";

                    label1.Text = dr["Question"].ToString() + "";
                    label2.Text = "Question" + " " + dr["QNo"].ToString() + "";
                    GlobalVar.QNo = Convert.ToInt32(dr["QNo"].ToString());
                }
                if (xName.Length > 0)
                {
                    ddlTestSheet.Text = xName;
                }
                conn.Close();
            }
        }

        private void GetData(int intCaseStudyId)
        {

            if (ddlTestSheet.Text == "")
            {
                MessageBox.Show("Please Select Test Sheet");
            }
            else
            {
                string str3 = "Select * from QuestionDetail where FK_CaseStudyId = " + intCaseStudyId + "";

                SqlDataAdapter sda = new SqlDataAdapter(str3, connectionString);
                dt = new System.Data.DataTable();
                sda.Fill(dt);
                if (currentRow < 0)
                {
                    return;
                }

                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        if (currentRow > dt.Rows.Count)
                        {
                            currentRow = 0;
                        }
                        if (currentRow < dt.Rows.Count)
                        {

                            QNo = Convert.ToInt32(dt.Rows[currentRow]["QNo"].ToString());
                            GlobalVar.QNo = QNo;
                            var row = dt.Rows[currentRow];
                            ddlTestSheet.Text = row["SheetName"].ToString();

                            label1.Text = row["Question"].ToString();
                            label2.Text = "Question" + " " + QNo;
                            Excel._Worksheet xlWorksheet = (Excel.Worksheet)xlApp.ActiveSheet;
                            Excel.Range xlRange = xlWorksheet.UsedRange;
                            if (row["SheetIndex"] + "" != string.Empty)
                            {
                                Worksheet sheet = (Worksheet)xlApp.Worksheets[Convert.ToInt32(row["SheetIndex"].ToString())];
                                sheet.Select(Type.Missing);
                            }
                            else
                            {
                                Worksheet sheet = (Worksheet)xlApp.Worksheets[1];
                                sheet.Select(Type.Missing);
                            }
                            currentRow++;
                            incrementPbar();

                        }
                    }
                }

            }

        }

        private void GetDataMinus(int intCaseStudyId)
        {

            if (ddlTestSheet.Text == "")
            {
                MessageBox.Show("Please Select Test Sheet");
            }
            else
            {
                string str3 = "Select * from QuestionDetail where FK_CaseStudyId = " + intCaseStudyId + "";

                SqlDataAdapter sda = new SqlDataAdapter(str3, connectionString);
                dt = new System.Data.DataTable();
                sda.Fill(dt);
                currentRow--;
                if (currentRow < 0)
                {
                    return;
                }

                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        if (currentRow > dt.Rows.Count)
                        {
                            currentRow = 0;
                        }
                        if (currentRow < dt.Rows.Count)
                        {

                            QNo = Convert.ToInt32(dt.Rows[currentRow]["QNo"].ToString());
                            var row = dt.Rows[currentRow];
                            ddlTestSheet.Text = row["SheetName"].ToString(); ;

                            label1.Text = row["Question"].ToString();
                            label2.Text = "Question" + " " + QNo;
                            Excel._Worksheet xlWorksheet = (Excel.Worksheet)xlApp.ActiveSheet;
                            Excel.Range xlRange = xlWorksheet.UsedRange;

                            if (row["SheetIndex"] + "" != string.Empty)
                            {
                                Worksheet sheet = (Worksheet)xlApp.Worksheets[Convert.ToInt32(row["SheetIndex"].ToString())];
                                sheet.Select(Type.Missing);
                            }
                            else
                            {
                                Worksheet sheet = (Worksheet)xlApp.Worksheets[1];
                                sheet.Select(Type.Missing);
                            }
                            incrementPbar();
                        }
                    }
                }

            }
        }
        
                public int CornerRadius { get; set; } = 30; // Corner radius

        protected override void OnPaint(PaintEventArgs pevent)
        {
            Graphics g = pevent.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // Create a rounded rectangle path
            GraphicsPath path = new GraphicsPath();
            path.AddArc(0, 0, CornerRadius, CornerRadius, 180, 90);
            path.AddArc(Width - CornerRadius, 0, CornerRadius, CornerRadius, 270, 90);
            path.AddArc(Width - CornerRadius, Height - CornerRadius, CornerRadius, CornerRadius, 0, 90);
            path.AddArc(0, Height - CornerRadius, CornerRadius, CornerRadius, 90, 90);
            path.CloseFigure();

            // Set the button's region
            this.Region = new Region(path);

            // Fill the button background
            using (SolidBrush brush = new SolidBrush(BackColor))
            {
                g.FillPath(brush, path);
            }

            // Draw the button text
            TextRenderer.DrawText(g, Text, Font, ClientRectangle, ForeColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            this.Invalidate(); // Redraw on resize
        }
    

        private void button1_Click(object sender, EventArgs e)

        { if (GlobalVar.intOpen == 1)
            {
                conn.Open();
                string selectquery = "SELECT * FROM QuestionDetail where QNo = " + GlobalVar.QNo + " and FK_CaseStudyId=" + GlobalVar.CaseStudyId + "";
                SqlCommand cmd = new SqlCommand(selectquery, conn);
                SqlDataReader reader1;
                reader1 = cmd.ExecuteReader();

                if (reader1.Read())
                {

                //    Form3 frm3 = new Form3();

                    label3.Visible = false;
                    label3.Text = reader1.GetValue(6).ToString();
                    GlobalVar.QLink = reader1.GetValue(12).ToString();
                    GlobalVar.Qhint = reader1.GetValue(7).ToString();
                    GlobalVar.frm3 = new Form3();
                    GlobalVar.frm3.Show();//   frm3.Show();
                }
                else
                {
                    MessageBox.Show("NO DATA FOUND");
                }
                conn.Close();
                GlobalVar.intOpen = GlobalVar.intOpen + 1;
            }

        }

        protected void CloseNext()
        {
            if (timer1.Enabled)
            {
                timer1.Stop();
            }
            else
            {
                timer1.Start();
            }

            DtFinal = DateTime.Now.TimeOfDay;
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }

            conn.Open();
            intCaseStudyId = GlobalVar.CaseStudyId;
            //SqlCommand cmd = new SqlCommand("select * from QuestionDetail where SheetName='"+ddlTestSheet.Text+"' and FK_CaseStudyId="+GlobalVar.CaseStudyId+"",conn);
            //SqlDataReader drs= cmd.ExecuteReader();
            //if(drs.HasRows)
            //{
            //    drs.Read();
            //    intCaseStudyId =Convert.ToInt32(drs["FK_CaseStudyId"].ToString());
            //}

            if (intCaseStudyId > 0)
            {
                //if (conn.State == ConnectionState.Open)
                //{
                //    conn.Close();
                //}
                //conn.Open();

                SqlDataAdapter daAns = new SqlDataAdapter("Select * from AnswerDetail where PK_AnswerId=0", conn);
                System.Data.DataTable dts = new System.Data.DataTable();

                daAns.Fill(dts);
                DataRow dr = dts.NewRow();
                dr["UserName"] = GlobalVar.GEmailId;
                dr["UserId"] = GlobalVar.GlobalUserId;
                dr["CaseStudyId"] = intCaseStudyId;
                dr["Question"] = label1.Text;
                dr["TimerStart"] = DtStart;
                dr["TimerStop"] = DtFinal;

                dts.Rows.Add(dr);
                SqlCommandBuilder cmb = new SqlCommandBuilder(daAns);
                daAns.Update(dt);





                //string query = "insert into AnswerDetail(UserName,UserId,CaseStudyId,Question,TimerStart,TimerStop) values('" + strUserName + "'," + intUserId + ",'" + intCaseStudyId + "','" + label1.Text + "','" + DtStart + "','" + DtFinal + "')";
                //SqlCommand cmd = new SqlCommand();
                //cmd = new SqlCommand(query, conn);
                //cmd.ExecuteNonQuery();
                try
                {

                    lblTimer.Text = "";
                    timer1.Dispose();
                    val = 0;
                    a = val++;
                    lblTimer.Text = a.ToString();

                    timer1.Start();
                    conn.Close();
                }

                catch (Exception ex)
                {
                    throw;
                }
                int rowCur = (currentRow + 1);

                GetData(intCaseStudyId);
            }
        }

        public void button3_Click(object sender, EventArgs e)
        {
            if (timer1.Enabled)
            {
                timer1.Stop();
            }
            else
            {
                timer1.Start();
            }

            DtFinal = DateTime.Now.TimeOfDay;
            if(conn.State==ConnectionState.Open)
            {
                conn.Close();
            }
            
            conn.Open();
            intCaseStudyId = GlobalVar.CaseStudyId;
            //SqlCommand cmd = new SqlCommand("select * from QuestionDetail where SheetName='"+ddlTestSheet.Text+"' and FK_CaseStudyId="+GlobalVar.CaseStudyId+"",conn);
            //SqlDataReader drs= cmd.ExecuteReader();
            //if(drs.HasRows)
            //{
            //    drs.Read();
            //    intCaseStudyId =Convert.ToInt32(drs["FK_CaseStudyId"].ToString());
            //}

            if (intCaseStudyId > 0)
            {
                //if (conn.State == ConnectionState.Open)
                //{
                //    conn.Close();
                //}
                //conn.Open();

                SqlDataAdapter daAns = new SqlDataAdapter("Select * from AnswerDetail where PK_AnswerId=0", conn);
                System.Data.DataTable dts = new System.Data.DataTable();

                daAns.Fill(dts);
                DataRow dr = dts.NewRow();
                dr["UserName"] = GlobalVar.GEmailId;
                dr["UserId"] = GlobalVar.GlobalUserId;
                dr["CaseStudyId"] = intCaseStudyId;
                dr["Question"] = label1.Text;
                dr["TimerStart"] = DtStart;
                dr["TimerStop"] = DtFinal;

                dts.Rows.Add(dr);
                SqlCommandBuilder cmb = new SqlCommandBuilder(daAns);
                daAns.Update(dt);





                //string query = "insert into AnswerDetail(UserName,UserId,CaseStudyId,Question,TimerStart,TimerStop) values('" + strUserName + "'," + intUserId + ",'" + intCaseStudyId + "','" + label1.Text + "','" + DtStart + "','" + DtFinal + "')";
                //SqlCommand cmd = new SqlCommand();
                //cmd = new SqlCommand(query, conn);
                //cmd.ExecuteNonQuery();
                try
                {

                    lblTimer.Text = "";
                    timer1.Dispose();
                    val = 0;
                    a = val++;
                    lblTimer.Text = a.ToString();

                    timer1.Start();
                    conn.Close();
                }

               catch (Exception ex)
                {
                    throw;
                }
                int rowCur = (currentRow + 1);

                GetData(intCaseStudyId);
            }
        }

        //public void loadFiles()
        //{

        //    string str3 = "Select * from QuestionDetail where SheetName = '" + ddlTestSheet.SelectedItem + "' and FK_CaseStudyId="+GlobalVar.CaseStudyId+"";
        //    SqlDataAdapter sda = new SqlDataAdapter(str3, connectionString);
        //    dt = new System.Data.DataTable();
        //    sda.Fill(dt);
        //    int v = dt.Rows.Count;
        //    progressBar1.Minimum = 1;
        //    progressBar1.Maximum = v;
        //    progressBar1.Value =1;
        //    progressBar1.Step = 1;
        //}

        //public void incrementPbar()
        //{
        //    progressBar1.PerformStep();
        //}

        //public void decrementPbar()
        //{
        //    if (progressBar1.Value > 1)
        //    {
        //        progressBar1.Value--;
        //    }
        //}

        // Variable to store the total number of rows
        private int totalRows = 0;

        public void loadFiles()
        {
            // SQL Query to get data
            string str3 = "Select * from QuestionDetail where SheetName = '" + ddlTestSheet.SelectedItem + "' and FK_CaseStudyId=" + GlobalVar.CaseStudyId;
            SqlDataAdapter sda = new SqlDataAdapter(str3, connectionString);
            dt = new System.Data.DataTable();
            sda.Fill(dt);

            // Get the total number of rows
            totalRows = dt.Rows.Count;

            // Set the progress bar's minimum and maximum values
            progressBar1.Minimum = 1;
            progressBar1.Maximum = totalRows;
            progressBar1.Value = 1;

            // Initialize Step to increment progress bar by 1
            progressBar1.Step = 1;

            // Initialize percentage label
            label5.Text = "0%";
        }

        public void incrementPbar()
        {
            // Increment the progress bar value by one step
            progressBar1.PerformStep();

            // Calculate the current percentage
            int currentPercentage = (int)((double)progressBar1.Value / totalRows * 100);

            // Update the percentage label
           label5.Text = $"{currentPercentage}%";
        }

        public void decrementPbar()
        {
            progressBar1.PerformStep();
            // Decrement progress bar value by one step, but not below 1
            if (progressBar1.Value > 1)
            {
                progressBar1.Value--;
            }

            // Calculate the current percentage
            int currentPercentage = (int)((double)progressBar1.Value / totalRows * 100);

            // Update the percentage label
            label5.Text = $"{currentPercentage}%";
        }


        private void button2_Click(object sender, EventArgs e)
        {
            string str3 = "Select * from QuestionDetail where SheetName = '" + this.ddlTestSheet.Text + "' and FK_CaseStudyId="+GlobalVar.CaseStudyId+"";
            SqlDataAdapter sda = new SqlDataAdapter(str3, connectionString);
            dt = new System.Data.DataTable();
            sda.Fill(dt);
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    if (currentRow >= 0)
                    {
                        if (currentRow < dt.Rows.Count)
                        {
                            QNo = Convert.ToInt32(dt.Rows[currentRow]["QNo"].ToString());
                            var row = dt.Rows[currentRow];
                            label1.Text = row["Question"].ToString();
                            label2.Text = "Question" + " " + QNo;
                            GlobalVar.QNo = QNo;
                            decrementPbar();
                        }
                    }
                }
            }


            if (currentRow > 0)
            {
                GetDataMinus(intCaseStudyId);
            }

        }
        static void DisplayListOfMacros(Workbook workbook)
        {
            foreach (VBComponent component in workbook.VBProject.VBComponents)
            {
                Console.WriteLine(component.Name);
                string selectedMacro = Console.ReadLine();
                RunMacro(workbook, selectedMacro);
            }
        }
        static void RunMacro(Workbook workbook, string macroName)
        {
            try
            {
                workbook.Application.Run(macroName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error running macro '{macroName}': {ex.Message}");
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Do you want to save all info...", "Sure", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                btnCloseExcel_Click(sender, e);

            MessageBox.Show("Scanning Your Answers...");
            button7_Click(sender,e);
            button5_Click(sender, e);
            button6_Click(sender, e);
            MessageBox.Show("Answers submitted successfully...");
                Form2 main = new Form2();
                main.ShowDialog()  ;
            }
        }

        private void ddlTestSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddlTestSheet.Text == "")
            {
                MessageBox.Show("Please Select CaseSheet");
                return;
            }
            else
            {
                DtStart = DateTime.Now.TimeOfDay;
                GlobalVar.strSheetName = ddlTestSheet.Text; 
                GetData(intCaseStudyId);
                loadFiles();
                val = 0;
                a = val++;
                lblTimer.Text = a.ToString();
                timer1.Start();
            }
            
        }

        //int sec = 0;
        int min = 00;
        int hour = 00;
        private void timer1_Tick(object sender, EventArgs e)
        {
            a = val++;
            lblTimer.Text = a.ToString();
            a++;
            if (a == 60)
            {
                min++;
                a = 00;
            }
            if (min == 60)
            {
                hour++;
                min = 00;
            }

            lblTimer.Text = hour.ToString() + " : " + min.ToString() + " : " + a.ToString();
        }
        private void Grid1Entry(int Qno)
        {
            if (ddlTestSheet.Text.Length <= 0)
            {
                MessageBox.Show("Please Select Case Study!");
                ddlTestSheet.Focus();
                return;
            }
            string connectionString = ConfigurationManager.ConnectionStrings["LMS_SQL"].ConnectionString;

            SqlConnection con = new SqlConnection(connectionString);
            string exportPath = "";
            string fullName = Assembly.GetEntryAssembly().Location;
            string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
            string myExcelPath = System.IO.Path.GetDirectoryName(fullName);

            string fname = myExcelPath + "\\Financial_Services_student_book.xlsx"; 
            FileInfo fileInfo = new FileInfo(fileExcel);
            fileInfo.IsReadOnly = false;

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[ddlTestSheet.Text];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            string EXPORT_TO_DIRECTORY = myExcelPath;
            exportPath = EXPORT_TO_DIRECTORY;

            Excel.Application app = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Microsoft.Office.Interop.Excel.Application;

            exportPath = EXPORT_TO_DIRECTORY;

            Excel.Workbook wb = app.ActiveWorkbook;

            SqlDataAdapter da = new SqlDataAdapter("select * from QuestionDetail where SheetName='" + ddlTestSheet.Text + "' and QNo="+Qno+" and FK_CaseStudyId="+GlobalVar.CaseStudyId+"", con);
            DataSet ds = new DataSet();
            da.Fill(ds);
            if (ds.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    if ((dr["QType"] + "").ToString().Trim() == "graph")
                    {
                        if (dr["QCell"] != null)
                        {
                            if (dr["QOrder"] != null)
                            {
                                string[] QNos = dr["QCell"].ToString().Split(',');
                                int QNo1 = Convert.ToInt32(QNos[0].ToString());
                                int QNo2 = Convert.ToInt32(QNos[1].ToString());
                                
                                string cc = xlRange.Cells[QNo1, QNo2].Value2.ToString();

                                if (Convert.ToInt32(cc) == Convert.ToInt32(dr["QNo"].ToString()))
                                {
                                    Excel.ChartObjects chartObjectsObj = (Excel.ChartObjects)(xlWorkbook.Sheets[ddlTestSheet.Text].ChartObjects(Type.Missing));

                                    if (chartObjectsObj.Count > 0)
                                    {
                                        foreach (ChartObject coObj in chartObjectsObj)
                                        {
                                            coObj.Select();
                                            Excel.Chart chartObj = (Excel.Chart)coObj.Chart;


                                            if (dr["TitleQ"].ToString().ToUpper() == chartObj.ChartTitle.Text.ToUpper())
                                            {
                                                chartObj.Export(exportPath + @"\" + chartObj.Name + ".bmp", "bmp", false);

                                                Image myImage = Image.FromFile(exportPath + @"\" + chartObj.Name + ".bmp");
                                                byte[] data;
                                                using (MemoryStream ms = new MemoryStream())
                                                {
                                                    myImage.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                                                    data = ms.ToArray();
                                                }

                                                dr["GraphImage1"] = data;
                                                SqlCommandBuilder cmb = new SqlCommandBuilder(da);
                                                da.Update(ds);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Charts Not Found!");
                                        return;
                                    }
                                }

                            }
                        }

                    }
                    else if ((dr["QType"] + "").ToString().Trim() == "pivot")
                    {
                        dataGridView1.ColumnCount = colCount;
                        dataGridView1.RowCount  = rowCount;


                        string[] QCell1 = dr["CellName"].ToString().Split(',');
                        string[] QCell2 = dr["CellNameTo"].ToString().Split(',');

                        int I1 = Convert.ToInt32(QCell1[0].ToString());
                        int J1 = Convert.ToInt32(QCell1[1].ToString());

                        int I2 = Convert.ToInt32(QCell2[0].ToString());
                        int J2 = Convert.ToInt32(QCell2[1].ToString());
                        for (int i = I1; i <= I2; i++)
                        {
                            for (int j = J1; j <= J2; j++)
                            {
                                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                                {
                                    dataGridView1.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Formula; // xlRange.Cells[i, j].Value2.ToString();
                                }
                            }
                        }
                    }

                    else
                    {
                       string[] QCell = dr["CellName"].ToString().Split(',');
                       dataGridView1.ColumnCount = colCount+1;
                       dataGridView1.RowCount = rowCount+1;

                        for (int i = 1; i <= rowCount; i++)
                        {
                            for (int j = 1; j <= colCount; j++)
                            {
                                if (Convert.ToInt32(QCell[0]) == i && Convert.ToInt32(QCell[1]) == j)
                                {
                                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                                    {

                                        dataGridView1.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Formula; // xlRange.Cells[i, j].Value2.ToString();
                                    }
                                }
                            }
                        }
                    }

                }

            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            Marshal.FinalReleaseComObject(xlApp);
            Marshal.FinalReleaseComObject(xlWorkbook);

            Marshal.FinalReleaseComObject(xlWorksheet);
        }
        private void Grid2Entry(int Qno)
        {
            string exportPath = "";
            string fullName = Assembly.GetEntryAssembly().Location;
            string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
            string myExcelPath = System.IO.Path.GetDirectoryName(fullName);
            string fname = myExcelPath + "\\Financial_Services.xlsx";

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[ddlTestSheet.Text];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            string EXPORT_TO_DIRECTORY = myExcelPath;
            exportPath = EXPORT_TO_DIRECTORY;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;


            Excel.Application app = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Microsoft.Office.Interop.Excel.Application;

            if (exportPath == "")
                exportPath = EXPORT_TO_DIRECTORY;

            Excel.Workbook wb = app.ActiveWorkbook;
       
            dataGridView2.ColumnCount = colCount;
            dataGridView2.RowCount = rowCount;
            string connectionString = ConfigurationManager.ConnectionStrings["LMS_SQL"].ConnectionString;

            SqlConnection con = new SqlConnection(connectionString);

            SqlDataAdapter da = new SqlDataAdapter("select * from QuestionDetail where SheetName='" + ddlTestSheet.Text + "' and QNo="+Qno+"", con);
            DataSet ds = new DataSet();
            da.Fill(ds);
            if (ds.Tables[0].Rows.Count > 0)
            {

                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    if ((dr["QType"] + "").ToString().Trim() == "graph")
                    {
                        if (dr["QCell"] != null)
                        {
                            if (dr["QOrder"] != null)
                            {
                                string[] QNos = dr["QCell"].ToString().Split(',');
                                int QNo1 = Convert.ToInt32(QNos[0].ToString());
                                int QNo2 = Convert.ToInt32(QNos[1].ToString());
                                string cc = xlRange.Cells[QNo1, QNo2].Value2.ToString();

                                if (Convert.ToInt32(cc) == Convert.ToInt32(dr["QNo"].ToString()))
                                {
                                    Excel.ChartObjects chartObjectsObj = (Excel.ChartObjects)(xlWorkbook.Sheets[ddlTestSheet.Text].ChartObjects(Type.Missing));
                                    if (chartObjectsObj.Count > 0)
                                    {
                                        foreach (ChartObject coObj in chartObjectsObj)
                                        {
                                            coObj.Select();
                                            Excel.Chart chartObj = (Excel.Chart)coObj.Chart;

                                            if (dr["TitleQ"].ToString().ToUpper() == chartObj.ChartTitle.Text.ToUpper())
                                            {
                                                chartObj.Export(exportPath + @"\ops.bmp", "bmp", false);
                                                _barray2 = ImageToBinary(exportPath + @"\ops.bmp");

                                                Image myImage = Image.FromFile(exportPath + @"\ops.bmp");

                                                byte[] data;
                                                using (MemoryStream ms = new MemoryStream())
                                                {
                                                    myImage.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                                                    data = ms.ToArray();
                                                }


                                                dr["GraphImage2"] = data;
                                                SqlCommandBuilder cmb = new SqlCommandBuilder(da);
                                                da.Update(ds);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Charts Not Found!");
                                        return;
                                    }
                                        
                                        
                                }
                            }
                        }
                    }
                    else if ((dr["QType"] + "").ToString().Trim() == "pivot")
                    {
                        string[] QCell1 = dr["CellName"].ToString().Split(',');
                        string[] QCell2 = dr["CellNameTo"].ToString().Split(',');

                        int I1 = Convert.ToInt32(QCell1[0].ToString());
                        int J1 = Convert.ToInt32(QCell1[1].ToString());

                        int I2 = Convert.ToInt32(QCell2[0].ToString());
                        int J2 = Convert.ToInt32(QCell2[1].ToString());
                        for (int i = I1; i <= I2; i++)
                        {
                            for (int j = J1; j <= J2; j++)
                            {
                                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                                {
                                    dataGridView2.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Formula; // xlRange.Cells[i, j].Value2.ToString();
                                }
                            }
                        }
                    }

                    else
                    {

                        string[] QCell = dr["CellName"].ToString().Split(',');

                        for (int i = 1; i <= rowCount; i++)
                        {
                            for (int j = 1; j <= colCount; j++)
                            {

                                if (Convert.ToInt32(QCell[0]) == i && Convert.ToInt32(QCell[1]) == j)
                                {
                                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                                    {
                                        string xy = xlRange.Cells[i, j].Formula;
                                        dataGridView2.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Formula;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
        private void Grid3Entry(int Qno)
        {
            try
            {
               var index = dataGridView3.Columns.Add("R", "Answer");
                var index2 = dataGridView3.Columns.Add("R1", "Status");
                var index1 = dataGridView3.Rows.Add();

                //Form2 frmform2 = new Form2();
               // frmform2.Close();
               
                int c = 2;
                int d = 2;
                
                dataGridView3.Rows.Clear();

                string connectionString = ConfigurationManager.ConnectionStrings["LMS_SQL"].ConnectionString;

                SqlConnection con = new SqlConnection(connectionString);
                string grid2 = "";
                string grid1 = "";
                Worksheet worksheet = xlWorkBook.Sheets[ddlTestSheet.Text];

                SqlDataAdapter da = new SqlDataAdapter("select * from QuestionDetail where  FK_CaseStudyId="+ GlobalVar.CaseStudyId + " and SheetName='" + ddlTestSheet.Text + "' and Qno="+Qno+"", con);
                DataSet ds = new DataSet();
                da.Fill(ds);
                if (ds.Tables[0].Rows.Count > 0)
                {

                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {

                        if ((dr["QType"] + "").ToString().Trim() == "graph")
                        {
                            if (dr["GraphImage1"] != null && dr["GraphImage2"] != null)
                            {
                                if (dr["GraphImage1"].ToString().Length > 0 && dr["GraphImage2"].ToString().Length > 0)
                                {

                                    _barray1 = (byte[])dr["GraphImage1"];
                                    _barray2 = (byte[])dr["GraphImage2"];

                                    Image img = ConvertByteArrayToImage(_barray1);  
                                    Bitmap bmp1 = new Bitmap(img);
                                    Image img1 = ConvertByteArrayToImage(_barray2); 
                                    Bitmap bmp2 = new Bitmap(img1);

                                    string fullName = Assembly.GetEntryAssembly().Location;
                                    string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
                                    string myExcelPath = System.IO.Path.GetDirectoryName(fullName);
                                    Excel.Application xlApp = new Excel.Application();

                                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileExcel);
                                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[ddlTestSheet.Text]; // Insert your sheet index here
                                    Excel.Range xlRange = xlWorksheet.UsedRange;

                                    xlApp.DisplayAlerts = false;
                                    Workbook workbook = xlApp.ActiveWorkbook;
                                    xt = Compare(bmp1, bmp2);

                                    if (xt == CompareResult.ciCompareOk)
                                    {
                                        if ((dr["AnswerCell"] + "").ToString().Trim().Length > 0)
                                        {
                                            string[] QCell = dr["AnswerCell"].ToString().Split(',');


                                            worksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1])] = "Correct";
                                            worksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1])].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);

                                            workbook.Save();
                                            xlApp.Workbooks.Close();
                                            xlApp.Quit();

                                            Marshal.ReleaseComObject(xlWorksheet);
                                            Marshal.ReleaseComObject(xlWorkbook);
                                            Marshal.ReleaseComObject(xlApp);
                                            dataGridView3.Rows[index1].Cells["R1"].Value = "Correct";

                                            AnsStatus = "Correct";
                                        }

                                        GlobalVariables.ModeType = 1;
                                        MsgBox msg = new MsgBox();
                                        msg.ShowDialog();
                                        if(GlobalVar.intClick==2)
                                        {
                                            saveAns();
                                            CloseNext();
                                        }
                                    }
                                    else  
                                    {

                                        GlobalVariables.ModeType = 2;
                                        MsgBox msg = new MsgBox();
                                        msg.ShowDialog();
                                        string[] QCell = dr["AnswerCell"].ToString().Split(',');

                                        worksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1]) + 1] = "Incorrect";
                                        worksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1]) + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                                        workbook.Save();
                                        xlApp.Workbooks.Close();
                                        xlApp.Quit();

                                        Marshal.ReleaseComObject(xlWorksheet);
                                        Marshal.ReleaseComObject(xlWorkbook);
                                        Marshal.ReleaseComObject(xlApp);
                                        dataGridView3.Rows[index1].Cells["R1"].Value = "Incorrect";

                                        AnsStatus = "Incorrect";
                                        if (GlobalVar.intClick == 2)
                                        {
                                            saveAns();
                                            CloseNext();
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Chart Not Found!");
                                    return;
                                }
                            }
                            else
                            {
                                MessageBox.Show("Chart Not Found!");
                                return;
                            }

                        }
                        else if ((dr["QType"] + "").ToString().Trim() == "pivot")
                        {
                            string fullName = Assembly.GetEntryAssembly().Location;
                            string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
                            string myExcelPath = System.IO.Path.GetDirectoryName(fullName);

                            Excel.Application xlApp = new Excel.Application();

                            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileExcel);
                            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[ddlTestSheet.Text]; 
                            Excel.Range xlRange = xlWorksheet.UsedRange;

                            string[] QCell1 = dr["CellName"].ToString().Split(',');
                            string[] QCell2 = dr["CellNameTo"].ToString().Split(',');

                            int I1 = Convert.ToInt32(QCell1[0].ToString());
                            int J1 = Convert.ToInt32(QCell1[1].ToString());

                            int I2 = Convert.ToInt32(QCell2[0].ToString());
                            int J2 = Convert.ToInt32(QCell2[1].ToString());
                            int int1 = 0;
                            int int2 = 0;
                            string str1 = "";
                            string str2 = "";


                            for (int i = I1; i <= I2; i++)
                            {
                                for (int j = J1; j <= J2; j++)
                                {

                                    if (dataGridView1.Rows[i - 1].Cells[j - 1].Value != null && dataGridView2.Rows[i - 1].Cells[j - 1].Value != null)
                                    {
                                        int1 = int1 + 1;
                                        str1 = dataGridView1.Rows[i - 1].Cells[j - 1].Value.ToString();
                                        str2 = dataGridView2.Rows[i - 1].Cells[j - 1].Value.ToString();

                                        if (str1 == str2)
                                        {
                                            int2 = int2 + 1;
                                        }
                                    }
                                }
                            }
                            if (int2 == int1)
                            {
                                string[] AnsCell = dr["AnswerCell"].ToString().Split(',');
                                if (AnsCell != null || AnsCell.Length > 1)
                                {

                                    worksheet.Cells[Convert.ToInt32(AnsCell[0]), Convert.ToInt32(AnsCell[1])] = "Correct";
                                    worksheet.Cells[Convert.ToInt32(AnsCell[0]), Convert.ToInt32(AnsCell[1])].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                                    GlobalVariables.ModeType = 1;
                                    AnsStatus = "Correct";
                                    MsgBox msgs = new MsgBox();
                                    msgs.ShowDialog();
                                    
                                    if (GlobalVar.intClick==2)
                                    {
                                        saveAns();
                                        CloseNext();
                                    }
                                }
                            }
                            else
                            {
                                string[] AnsCell = dr["AnswerCell"].ToString().Split(',');
                                if (AnsCell != null || AnsCell.Length > 1)
                                {
                                    worksheet.Cells[Convert.ToInt32(AnsCell[0]), Convert.ToInt32(AnsCell[1])] = "Incorrect";
                                    worksheet.Cells[Convert.ToInt32(AnsCell[0]), Convert.ToInt32(AnsCell[1])].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                                    GlobalVariables.ModeType = 2;
                                    AnsStatus = "Incorrect";
                                      MsgBox msgaa = new MsgBox();
                                    msgaa.ShowDialog();
                                    if (GlobalVar.intClick == 2)
                                    {
                                        saveAns();
                                        CloseNext();
                                    }

                                }

                            }
                            xlApp.DisplayAlerts = false;
                            xlWorkBook.Save();
                            xlApp.Workbooks.Close();
                            xlApp.Quit();

                            Marshal.ReleaseComObject(xlWorksheet);
                            Marshal.ReleaseComObject(xlWorkbook);
                            Marshal.ReleaseComObject(xlApp);
                            dataGridView3.Rows[index1].Cells["R1"].Value = "Correct";

                        }

                        else
                        {

                            string[] QCell = dr["CellName"].ToString().Split(',');
                            int j = 0;
                            for (int i = 1; i < dataGridView1.RowCount; i++)
                            {
                                if (Convert.ToInt32(QCell[0]) == i)
                                {
                                    j = Convert.ToInt32(QCell[1]);

                                    if (dataGridView2.Rows[i - 1].Cells[Convert.ToInt32(QCell[1]) - 1].Value != null)
                                    {
                                        grid2 = dataGridView2.Rows[i - 1].Cells[Convert.ToInt32(QCell[1]) - 1].Value.ToString() + "";
                                    }
                                    if (dataGridView1.Rows[i - 1].Cells[Convert.ToInt32(QCell[1]) - 1].Value != null)
                                    {
                                        grid1 = dataGridView1.Rows[i - 1].Cells[Convert.ToInt32(QCell[1]) - 1].Value.ToString() + "";
                                    }

                                    index = dataGridView3.Columns.Add("R", "Answer");
                                    index2 = dataGridView3.Columns.Add("R1", "Status");
                                    index1 = dataGridView3.Rows.Add();
                                    dataGridView3.Rows[index1].Cells["R"].Value = grid1;

                                    if (grid1 == grid2)
                                    {
                                        Excel.Application xlApp = new Excel.Application();

                                        string fullName = Assembly.GetEntryAssembly().Location;
                                        string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
                                        string myExcelPath = System.IO.Path.GetDirectoryName(fullName);
                                        Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileExcel);
                                        Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[ddlTestSheet.Text]; // Insert your sheet index here
                                        Excel.Range xlRange = xlWorksheet.UsedRange;

                                        xlApp.DisplayAlerts = false;
                                        string[] AnsCell = dr["AnswerCell"].ToString().Split(',');
                                        if (AnsCell != null || AnsCell.Length > 1)
                                        {
                                            worksheet.Cells[Convert.ToInt32(AnsCell[0]), Convert.ToInt32(AnsCell[1])] = "Correct";
                                            worksheet.Cells[Convert.ToInt32(AnsCell[0]), Convert.ToInt32(AnsCell[1])].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                                        }
                                        else
                                        {
                                            worksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1]) + 1] = "Correct";
                                            worksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1]) + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);

                                        }
                                        AnsStatus = "Correct";
                                        xlWorkBook.Save();
                                        xlApp.Workbooks.Close();
                                        xlApp.Quit();

                                        Marshal.ReleaseComObject(xlWorksheet);
                                        Marshal.ReleaseComObject(xlWorkbook);
                                        Marshal.ReleaseComObject(xlApp);
                                  
                                        dataGridView3.Rows[index1].Cells["R1"].Value = "Correct";
                                        GlobalVariables.ModeType = 1;
                                        MsgBox msg = new MsgBox();
                                        msg.ShowDialog();
                                        if (GlobalVar.intClick==2)
                                        {
                                            saveAns();
                                            CloseNext();
                                           
                                        }
                                    }
                                    else
                                    {
                                        Excel.Application xlApp = new Excel.Application();

                                        string fullName = Assembly.GetEntryAssembly().Location;
                                        string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
                                        string myExcelPath = System.IO.Path.GetDirectoryName(fullName);

                                        fileExcel = myExcelPath + "\\Financial_Services_student_book.xlsx"; 
                                        FileInfo fileInfo = new FileInfo(fileExcel);
                                        fileInfo.IsReadOnly = false;

                                        Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileExcel);
                                        Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[ddlTestSheet.Text]; 
                                        Excel.Range xlRange = xlWorksheet.UsedRange;

                                         xlWorksheet = xlWorkbook.Sheets[ddlTestSheet.Text];

                                        string[] AnsCells = dr["AnswerCell"].ToString().Split(',');

                                        xlApp.DisplayAlerts = false;
                                        if (AnsCells.Length > 1)
                                        {
                                            worksheet.Cells[Convert.ToInt32(AnsCells[0]), Convert.ToInt32(AnsCells[1])] = "Incorrect";
                                            worksheet.Cells[Convert.ToInt32(AnsCells[0]), Convert.ToInt32(AnsCells[1])].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                        }
                                        else
                                        {
                                            worksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1]) + 1] = "Incorrect";
                                            worksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1]) + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                                        }
                                        xlWorkBook.Save();

                                        xlApp.Workbooks.Close();
                                        xlApp.Quit();

                                        Marshal.ReleaseComObject(xlWorksheet);
                                        Marshal.ReleaseComObject(xlWorkbook);
                                        Marshal.ReleaseComObject(xlApp);

                                        dataGridView3.Rows[index1].Cells["R1"].Value = "Incorrect";
                                  
                                        GlobalVariables.ModeType = 2;
                                        MsgBox msg = new MsgBox();
                                        msg.ShowDialog();
                                        if (GlobalVar.intClick == 2)
                                        {
                                            saveAns();
                                            CloseNext();

                                        }
                                    }

                                    index++;
                                    index1++;
                                    index2++;
                                    c++;

                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please Press Control+S to save and close...");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if(ddlTestSheet.Text.Length<=0)
            {
                MessageBox.Show("Please Select Case Study!");
                ddlTestSheet.Focus(); 
                return;
            }
            string connectionString = ConfigurationManager.ConnectionStrings["LMS_SQL"].ConnectionString;

            SqlConnection con = new SqlConnection(connectionString);
            string exportPath = "";
            string fullName = Assembly.GetEntryAssembly().Location;
            string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
            string myExcelPath = System.IO.Path.GetDirectoryName(fullName);

            string fname = myExcelPath + "\\Financial_Services_student_book.xlsx"; 
            FileInfo fileInfo = new FileInfo(fileExcel);
            fileInfo.IsReadOnly = false;
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[ddlTestSheet.Text];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            string EXPORT_TO_DIRECTORY = myExcelPath;
            exportPath = EXPORT_TO_DIRECTORY;

            Excel.Application app = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Microsoft.Office.Interop.Excel.Application;

            exportPath = EXPORT_TO_DIRECTORY;

            Excel.Workbook wb = app.ActiveWorkbook;



            SqlDataAdapter da = new SqlDataAdapter("select * from QuestionDetail where FK_CaseStudyId="+GlobalVar.CaseStudyId+" and SheetName='" + ddlTestSheet.Text + "'", con);
            DataSet ds = new DataSet();
            da.Fill(ds);
            if (ds.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    if ((dr["QType"] + "").ToString().Trim() == "graph")
                    {
                        if (dr["QCell"] != null)
                        {
                            if (dr["QOrder"] != null)
                            {
                                string[] QNos = dr["QCell"].ToString().Split(',');
                                int QNo1 = Convert.ToInt32(QNos[0].ToString());
                                int QNo2 = Convert.ToInt32(QNos[1].ToString());
                                string cc = xlRange.Cells[QNo1, QNo2].Value2.ToString();

                                if (Convert.ToInt32(cc) == Convert.ToInt32(dr["QNo"].ToString()))
                                {
                                    Excel.ChartObjects chartObjectsObj = (Excel.ChartObjects)(xlWorkbook.Sheets[ddlTestSheet.Text].ChartObjects(Type.Missing));

                                    Excel.ChartObject coObj = chartObjectsObj.Item(Convert.ToInt32(dr["QOrder"].ToString()));
                                    coObj.Select();
                                    Excel.Chart chartObj = (Excel.Chart)coObj.Chart;
                                    chartObj.Export(exportPath + @"\" + chartObj.Name + ".bmp", "bmp", false);
                       
                                    Image myImage = Image.FromFile(exportPath + @"\" + chartObj.Name + ".bmp");
                        
                                    byte[] data;
                                    using (MemoryStream ms = new MemoryStream())
                                    {
                                        myImage.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                                        data = ms.ToArray();
                                    }

                                    dr["GraphImage1"] = data;
                                    SqlCommandBuilder cmb = new SqlCommandBuilder(da);
                                    da.Update(ds);
                                }
                            }
                        }

                    }
                    else if ((dr["QType"] + "").ToString().Trim() == "pivot")
                        {
                        string[] QCell1= dr["CellName"].ToString().Split(',');
                        string[] QCell2 = dr["CellNameTo"].ToString().Split(',');
                        
                        int I1 = Convert.ToInt32(QCell1[0].ToString());
                        int J1 = Convert.ToInt32(QCell1[1].ToString());
                        
                        int I2 = Convert.ToInt32(QCell2[0].ToString());
                        int J2 = Convert.ToInt32(QCell2[1].ToString());
                        for (int i = I1; i <= I2; i++)
                        {
                            for (int j = J1; j <= J2; j++)
                            {
                                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                                    {
                                        dataGridView1.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Formula; // xlRange.Cells[i, j].Value2.ToString();
                                    }
               
                            }
                        }
                    }

                    else
                        {
                            string[] QCell = dr["CellName"].ToString().Split(',');
                            dataGridView1.ColumnCount = colCount;
                            dataGridView1.RowCount = rowCount;

                            for (int i = 1; i <= rowCount; i++)
                            {
                                for (int j = 1; j <= colCount; j++)
                                {
                                    if (Convert.ToInt32(QCell[0]) == i && Convert.ToInt32(QCell[1]) == j)
                                    {
                                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                                        {
                        
                                            dataGridView1.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Formula; 
                                        }
                                    }
                                }
                            }
                        }

                    }

                }

             
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            Marshal.FinalReleaseComObject(xlApp);
            Marshal.FinalReleaseComObject(xlWorkbook);

            Marshal.FinalReleaseComObject(xlWorksheet);
        }
        private void button5_Click(object sender, EventArgs e)
        {
            string exportPath="";
            string fullName = Assembly.GetEntryAssembly().Location;
            string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
            string myExcelPath = System.IO.Path.GetDirectoryName(fullName);

            string fname = myExcelPath + "\\Financial_Services.xlsx"; 


            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[ddlTestSheet.Text];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            string EXPORT_TO_DIRECTORY = myExcelPath;
            exportPath = EXPORT_TO_DIRECTORY;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;


            Excel.Application app = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Microsoft.Office.Interop.Excel.Application;

            if (exportPath == "")
                exportPath = EXPORT_TO_DIRECTORY;

            Excel.Workbook wb = app.ActiveWorkbook;

            dataGridView2.ColumnCount = colCount;
            dataGridView2.RowCount = rowCount;
            string connectionString = ConfigurationManager.ConnectionStrings["LMS_SQL"].ConnectionString;

            SqlConnection con = new SqlConnection(connectionString);

            SqlDataAdapter da = new SqlDataAdapter("select * from QuestionDetail where FK_CaseStudyId="+GlobalVar.CaseStudyId +" and SheetName='" + ddlTestSheet.Text + "'", con);
            DataSet ds = new DataSet();
            da.Fill(ds);
            if (ds.Tables[0].Rows.Count > 0)
            {

                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    if ((dr["QType"] + "").ToString().Trim() == "graph")
                    {
                        if (dr["QCell"] != null)
                        {
                            if (dr["QOrder"] != null)
                            {
                                string[] QNos = dr["QCell"].ToString().Split(',');
                                int QNo1 = Convert.ToInt32(QNos[0].ToString());
                                int QNo2 = Convert.ToInt32(QNos[1].ToString());
                                string cc = xlRange.Cells[QNo1, QNo2].Value2.ToString();

                                if (Convert.ToInt32(cc) == Convert.ToInt32(dr["QNo"].ToString()))
                                {
                                    Excel.ChartObjects chartObjectsObj = (Excel.ChartObjects)(xlWorkbook.Sheets[ddlTestSheet.Text].ChartObjects(Type.Missing));

                                    Excel.ChartObject coObj = chartObjectsObj.Item(Convert.ToInt32(dr["QOrder"].ToString()));
                                    coObj.Select();
                                    Excel.Chart chartObj = (Excel.Chart)coObj.Chart;
                                     chartObj.Export(exportPath + @"\ops.bmp", "bmp", false);
                                     _barray2 = ImageToBinary(exportPath + @"\ops.bmp");

                                    Image myImage = Image.FromFile(exportPath + @"\ops.bmp");

                                    byte[] data;
                                    using (MemoryStream ms = new MemoryStream())
                                    {
                                        myImage.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                                        data = ms.ToArray();
                                    }


                                    dr["GraphImage2"] = data;
                                    SqlCommandBuilder cmb = new SqlCommandBuilder(da);
                                    da.Update(ds);
                                }
                            }
                        }


                    }
                    else if ((dr["QType"] + "").ToString().Trim() == "pivot")
                    {
                        string[] QCell1 = dr["CellName"].ToString().Split(',');
                        string[] QCell2 = dr["CellNameTo"].ToString().Split(',');

                        int I1 = Convert.ToInt32(QCell1[0].ToString());
                        int J1 = Convert.ToInt32(QCell1[1].ToString());

                        int I2 = Convert.ToInt32(QCell2[0].ToString());
                        int J2 = Convert.ToInt32(QCell2[1].ToString());
                        for (int i = I1; i <= I2; i++)
                        {
                            for (int j = J1; j <= J2; j++)
                            {
                                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                                {
                                    dataGridView2.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Formula; // xlRange.Cells[i, j].Value2.ToString();
                                }
                            }
                        }
                    }

                    else
                    {

                        string[] QCell = dr["CellName"].ToString().Split(',');

                        for (int i = 1; i <= rowCount; i++)
                        {
                            for (int j = 1; j <= colCount; j++)
                            {

                                if (Convert.ToInt32(QCell[0]) == i && Convert.ToInt32(QCell[1]) == j)
                                {
                                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                                    {
                                        string xy = xlRange.Cells[i, j].Formula;
                                        dataGridView2.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Formula;

                                    }
                                }
                            }
                        }
                    }
                }
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
        public Image ConvertByteArrayToImage(byte[] byteArray)
        {
            MemoryStream ms = new MemoryStream(byteArray, 0, byteArray.Length);
            ms.Write(byteArray, 0, byteArray.Length);
            Image image = Image.FromStream(ms, true);
            return image;
        }
        
        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                var index = dataGridView3.Columns.Add("R", "Answer");
                var index2 = dataGridView3.Columns.Add("R1", "Status");
                var index1 = dataGridView3.Rows.Add();

                Form2 frmform2 = new Form2();
                frmform2.Close();
                int c = 2;
                int d = 2;
                dataGridView3.Rows.Clear();

                string connectionString = ConfigurationManager.ConnectionStrings["LMS_SQL"].ConnectionString;

                SqlConnection con = new SqlConnection(connectionString);
                string grid2 = "";
                string grid1 = "";

                SqlDataAdapter da = new SqlDataAdapter("select * from QuestionDetail  where FK_CaseStudyId="+GlobalVar.CaseStudyId+" and SheetName='" + ddlTestSheet.Text + "'", con);
                DataSet ds = new DataSet();
                da.Fill(ds);
                if (ds.Tables[0].Rows.Count > 0)
                {

                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {

                        if ((dr["QType"] + "").ToString().Trim() == "graph")
                        {
                            _barray1 = (byte[])dr["GraphImage1"];
                            _barray2 = (byte[])dr["GraphImage2"];

                            ImageConverter ic = new ImageConverter();
                            Image img = (Image)ic.ConvertFrom(_barray1);
                            Bitmap bmp1 = new Bitmap(img);
                            Image img1 = (Image)ic.ConvertFrom(_barray2);
                            Bitmap bmp2 = new Bitmap(img1);

                            string fullName = Assembly.GetEntryAssembly().Location;
                            string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
                            string myExcelPath = System.IO.Path.GetDirectoryName(fullName);

                            fileExcel = myExcelPath + "\\Financial_Services_student_book.xlsx";
                            FileInfo fileInfo = new FileInfo(fileExcel);
                            fileInfo.IsReadOnly = false;
                            Excel.Application xlApp = new Excel.Application();

                            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileExcel);
                            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[ddlTestSheet.Text]; 
                            Excel.Range xlRange = xlWorksheet.UsedRange;

                            xlApp.DisplayAlerts = false;

                            xt = Compare(bmp1, bmp2);
                                if(xt == CompareResult.ciCompareOk)
                              {
                                if((dr["AnswerCell"] + "").ToString().Trim().Length>0)
                                {
                                    string[] QCell = dr["AnswerCell"].ToString().Split(',');


                                    xlWorksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1])] = "Correct";
                                    xlWorksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1])].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);

                                    xlWorksheet.SaveAs(fileExcel);
                                    xlApp.Workbooks.Close();
                                    xlApp.Quit();

                                    Marshal.ReleaseComObject(xlWorksheet);
                                    Marshal.ReleaseComObject(xlWorkbook);
                                    Marshal.ReleaseComObject(xlApp);
                                    dataGridView3.Rows[index1].Cells["R1"].Value = "correct";
                               }


                                GlobalVariables.ModeType = 1;
                               
                                MsgBox msg = new MsgBox();
                                msg.ShowDialog();
                                if (GlobalVar.intClick == 2)
                                {

                                }

                            }
                            else  
                            {

                                GlobalVariables.ModeType = 2;
                                MsgBox msg = new MsgBox();
                                msg.ShowDialog();

                                string[] QCell = dr["AnswerCell"].ToString().Split(',');

                                xlWorksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1]) + 1] = "Incorrect";
                                xlWorksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1]) + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                                xlWorksheet.SaveAs(fileExcel);
                                xlApp.Workbooks.Close();
                                xlApp.Quit();

                                Marshal.ReleaseComObject(xlWorksheet);
                                Marshal.ReleaseComObject(xlWorkbook);
                                Marshal.ReleaseComObject(xlApp);
                                dataGridView3.Rows[index1].Cells["R1"].Value = "Incorrect";

                                if (GlobalVar.intClick == 2)
                                {
                                    CloseNext();
                                }
                            }

                        }
                        else if ((dr["QType"] + "").ToString().Trim() == "pivot")
                        {
                            string fullName = Assembly.GetEntryAssembly().Location;
                            string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
                            string myExcelPath = System.IO.Path.GetDirectoryName(fullName);

                            fileExcel = myExcelPath + "\\Financial_Services_student_book.xlsx"; 
                            FileInfo fileInfo = new FileInfo(fileExcel);
                            fileInfo.IsReadOnly = false;
                            Excel.Application xlApp = new Excel.Application();

                            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileExcel);
                            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[ddlTestSheet.Text];
                            Excel.Range xlRange = xlWorksheet.UsedRange;


                            string[] QCell1 = dr["CellName"].ToString().Split(',');
                            string[] QCell2 = dr["CellNameTo"].ToString().Split(',');

                            int I1 = Convert.ToInt32(QCell1[0].ToString());
                            int J1 = Convert.ToInt32(QCell1[1].ToString());

                            int I2 = Convert.ToInt32(QCell2[0].ToString());
                            int J2 = Convert.ToInt32(QCell2[1].ToString());
                            int int1 = 0;
                            int int2 = 0;
                            string str1 = "";
                            string str2 = "";


                            for (int i = I1; i <= I2; i++)
                            {
                                for (int j = J1; j <= J2; j++)
                                {
              
                                    if (dataGridView1.Rows[i - 1].Cells[j - 1].Value != null && dataGridView2.Rows[i - 1].Cells[j - 1].Value !=null)
                                    {
                                        int1 = int1 + 1;
                                        str1 = dataGridView1.Rows[i - 1].Cells[j - 1].Value.ToString();
                                        str2 = dataGridView2.Rows[i - 1].Cells[j - 1].Value.ToString();

                                        if (str1 == str2)
                                        {
                                            int2 = int2 + 1;
                                        }
                                    }
                                   


                                }
                            }
                            if (int2 == int1)
                            {
                                string[] AnsCell = dr["AnswerCell"].ToString().Split(',');
                                if (AnsCell != null || AnsCell.Length > 1)
                                {

                                    xlWorksheet.Cells[Convert.ToInt32(AnsCell[0]), Convert.ToInt32(AnsCell[1])] = "correct";
                                    xlWorksheet.Cells[Convert.ToInt32(AnsCell[0]), Convert.ToInt32(AnsCell[1])].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                                }
                                MessageBox.Show(dr["QNo"].ToString());
                                GlobalVariables.ModeType = 1;
                                MsgBox msgss = new MsgBox();
                                msgss.ShowDialog();
                                if (GlobalVar.intClick == 2)
                                {

                                }
                            }
                            else
                            {
                                string[] AnsCell = dr["AnswerCell"].ToString().Split(',');
                                if (AnsCell != null || AnsCell.Length > 1)
                                {

                                    xlWorksheet.Cells[Convert.ToInt32(AnsCell[0]), Convert.ToInt32(AnsCell[1])] = "Incorrect";
                                    xlWorksheet.Cells[Convert.ToInt32(AnsCell[0]), Convert.ToInt32(AnsCell[1])].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                    GlobalVariables.ModeType = 2;
                                    MsgBox msgo = new MsgBox();
                                    msgo.ShowDialog();
                                    if (GlobalVar.intClick == 2)
                                    {
                                        CloseNext();
                                    }
                                }

                            }
                            xlApp.DisplayAlerts = false;
                            xlWorksheet.SaveAs(fileExcel);
                            xlApp.Workbooks.Close();
                            xlApp.Quit();

                            Marshal.ReleaseComObject(xlWorksheet);
                            Marshal.ReleaseComObject(xlWorkbook);
                            Marshal.ReleaseComObject(xlApp);
                            dataGridView3.Rows[index1].Cells["R1"].Value = "correct";

                            GlobalVariables.ModeType = 2;
                            MsgBox msg = new MsgBox();
                            msg.ShowDialog();
                            if (GlobalVar.intClick == 2)
                            {
                                CloseNext();
                            }
                        }


                        else
                        {

                            string[] QCell = dr["CellName"].ToString().Split(',');

                  
                            for (int i = 1; i < dataGridView1.RowCount; i++)
                            {
                                if (Convert.ToInt32(QCell[0]) == i)
                                {
                     
                                    if (dataGridView2.Rows[i - 1].Cells[Convert.ToInt32(QCell[1]) - 1].Value != null)
                                    {
                                        grid2 = dataGridView2.Rows[i - 1].Cells[Convert.ToInt32(QCell[1]) - 1].Value.ToString() + "";
                                    }
                                    if (dataGridView1.Rows[i - 1].Cells[Convert.ToInt32(QCell[1]) - 1].Value != null)
                                    {
                                        grid1 = dataGridView1.Rows[i - 1].Cells[Convert.ToInt32(QCell[1]) - 1].Value.ToString() + "";
                                    }

                                    index = dataGridView3.Columns.Add("R", "Answer");
                                    index2 = dataGridView3.Columns.Add("R1", "Status");
                                    index1 = dataGridView3.Rows.Add();
                                    dataGridView3.Rows[index1].Cells["R"].Value = grid1;

                                    if (grid1 == grid2)
                                    {
                                        Excel.Application xlApp = new Excel.Application();

                                        string fullName = Assembly.GetEntryAssembly().Location;
                                        string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
                                        string myExcelPath = System.IO.Path.GetDirectoryName(fullName);

                                        fileExcel = myExcelPath + "\\Financial_Services_student_book.xlsx"; //"D:\\CharakPoints.xlsx";
                                        FileInfo fileInfo = new FileInfo(fileExcel);
                                        fileInfo.IsReadOnly = false;

                                        Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileExcel);
                                        Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[ddlTestSheet.Text]; // Insert your sheet index here
                                        Excel.Range xlRange = xlWorksheet.UsedRange;

                                        xlApp.DisplayAlerts = false;

                                        string[] AnsCell  = dr["AnswerCell"].ToString().Split(',');
                                        if (AnsCell !=null || AnsCell.Length >1)
                                        {

                                            xlWorksheet.Cells[Convert.ToInt32(AnsCell[0]), Convert.ToInt32(AnsCell[1])] = "correct";
                                            xlWorksheet.Cells[Convert.ToInt32(AnsCell[0]), Convert.ToInt32(AnsCell[1])].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                                            GlobalVariables.ModeType = 1;
                                            MsgBox msg = new MsgBox();
                                            msg.ShowDialog();
                                            if (GlobalVar.intClick == 2)
                                            {

                                            }
                                        }
                                        else
                                        {
                                            xlWorksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1]) + 1] = "correct";
                                            xlWorksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1]) + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);

                                            GlobalVariables.ModeType = 1;
                                            MsgBox msg = new MsgBox();
                                            msg.ShowDialog();
                                            if (GlobalVar.intClick == 2)
                                            {

                                            }
                                        }
                                        xlWorksheet.SaveAs(fileExcel);
                                        xlApp.Workbooks.Close();
                                        xlApp.Quit();

                                        Marshal.ReleaseComObject(xlWorksheet);
                                        Marshal.ReleaseComObject(xlWorkbook);
                                        Marshal.ReleaseComObject(xlApp);
                                        dataGridView3.Rows[index1].Cells["R1"].Value = "correct";

                                    }
                                    else
                                    {
                                        Excel.Application xlApp = new Excel.Application();

                                        string fullName = Assembly.GetEntryAssembly().Location;
                                        string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
                                        string myExcelPath = System.IO.Path.GetDirectoryName(fullName);

                                        fileExcel = myExcelPath + "\\Financial_Services_student_book.xlsx"; //"D:\\CharakPoints.xlsx";
                                        FileInfo fileInfo = new FileInfo(fileExcel);
                                        fileInfo.IsReadOnly = false;

                                        Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileExcel);
                                        Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[ddlTestSheet.Text]; // Insert your sheet index here
                                        Excel.Range xlRange = xlWorksheet.UsedRange;

                                        string[] AnsCells = dr["AnswerCell"].ToString().Split(',');

                                        xlApp.DisplayAlerts = false;
                                        if (AnsCells.Length > 1)
                                        {
                                            xlWorksheet.Cells[Convert.ToInt32(AnsCells[0]), Convert.ToInt32(AnsCells[1])] = "incorrect";
                                            xlWorksheet.Cells[Convert.ToInt32(AnsCells[0]), Convert.ToInt32(AnsCells[1])].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                        }
                                        else
                                        {
                                            xlWorksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1]) + 1] = "incorrect";
                                            xlWorksheet.Cells[Convert.ToInt32(QCell[0]), Convert.ToInt32(QCell[1]) + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                                        }

                                        //xlWorksheet.Cells[1, 2] = "Value";
                                        xlWorksheet.SaveAs(fileExcel);
                                        xlApp.Workbooks.Close();
                                        xlApp.Quit();

                                        Marshal.ReleaseComObject(xlWorksheet);
                                        Marshal.ReleaseComObject(xlWorkbook);
                                        Marshal.ReleaseComObject(xlApp);

                                        dataGridView3.Rows[index1].Cells["R1"].Value = "incorrect";
                                        GlobalVariables.ModeType = 2;
                                        MsgBox msg = new MsgBox();
                                        msg.ShowDialog();

                                    }

                                    index++;
                                    index1++;
                                    index2++;
                                    c++;

                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please Press Control+S to save and close...");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            btnCloseExcel_Click(sender, e);
            GlobalVar.frm3.Close(); 
            string sourcePath = @myExcelPath;
            string targetPath = @myExcelPath;

            string sourceFile  = System.IO.Path.Combine(targetPath, SingleFileName);
            string destFile = System.IO.Path.Combine(sourcePath, GlobalVar.GlobalUserId + "_"+ SingleFileName);

            System.IO.File.Copy(sourceFile, destFile, true);


            string FullFileName = GlobalVar.GlobalUserId + "_" + SingleFileName;
            fileExcel = @myExcelPath+"\\"+ GlobalVar.GlobalUserId + "_" + SingleFileName;

            WebClient client = new WebClient();
            
            string myNames = System.IO.Path.GetFileName(fileExcel);
            string uploadWebUrl = "https://www.intallium.com/upload.aspx?Id="+ FullFileName + "";
            client.UploadFile(uploadWebUrl, fileExcel);
            MessageBox.Show("File Uploaded Successfully...");
            this.Close();
        }

        private void button9_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void btnCloseExcel_Click(object sender, EventArgs e)
        {
           foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                process.Kill();
            }
        }

        private void btnReadXml_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            DataSet ds1 = new DataSet();
            DataSet ds2 = new DataSet();
            DataSet ds3 = new DataSet();


            ds.ReadXml("D:\\sheet1.xml");

            XmlDocument doc = new XmlDocument();
            doc.Load("D:\\chart\\chart1.xml");

            foreach (XmlNode node in doc.DocumentElement.ChildNodes)
            {
                foreach (XmlNode locNode in node)
                {
                    if (locNode.Name == "loc")
                    {
                        string loc = locNode.InnerText;
                        Console.WriteLine(loc + Environment.NewLine);
                    }
                }
            }
        }
        public byte[] ImageToByteArray(System.Drawing.Image imageIn)
        {
            using (var ms = new MemoryStream())
            {
                imageIn.Save(ms, imageIn.RawFormat);
                return ms.ToArray();
            }
        }
        public static byte[] ImageToBinary(string imagePath)
        {
            FileStream fileStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
            byte[] buffer = new byte[fileStream.Length];
            fileStream.Read(buffer, 0, (int)fileStream.Length);
            fileStream.Close();
            return buffer;
        }
        private void btnExportImage_Click(object sender, EventArgs e)
        {
            _barray2 = ImageToBinary(@"F:\OPSS\Charts\Alpha Chart 1.jpg");
            _barray1 = ImageToBinary(@"F:\OPSS\Charts\AlphaChart1.jpg");

            ImageConverter ic = new ImageConverter();
            Image img = (Image)ic.ConvertFrom(_barray1);
            Bitmap bmp1 = new Bitmap(img);
            Image img1 = (Image)ic.ConvertFrom(_barray2);
            Bitmap bmp2 = new Bitmap(img1);
            if (Class1.Compare(bmp1, bmp2) == Class1.CompareResult.ciCompareOk)
            {
                MessageBox.Show("Images Are Same");
            }
            else if (Class1.Compare(bmp1, bmp2) == Class1.CompareResult.ciPixelMismatch)
            {
                MessageBox.Show("Images not matching");
            }
            else if (Class1.Compare(bmp1, bmp2) == Class1.CompareResult.ciSizeMismatch)
            {
                MessageBox.Show("Size Images not same");
            }

        }
        protected void saveAns()
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
            conn.Open();

            string ssql = "delete from UserAnswer  where UserId=" + GlobalVar.GlobalUserId + " and FK_CaseStudyId='" + GlobalVar.CaseStudyId + "' and QNo=" + GlobalVar.QNo + "";
            SqlCommand cmd = new SqlCommand(ssql, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
            //conn.Open();

            //ssql = "Insert into UserAnswer(UserId,EmailId,SheetName,QNo,FK_CaseStudyId,Status,Question) values(" + GlobalVar.GlobalUserId + ",'" + strEmailId + "','" + GlobalVar.strSheetName + "'," + QNo + "," + GlobalVar.CaseStudyId + ",'"+AnsStatus+"','"+label1.Text+"')";
            //cmd = new SqlCommand(ssql, conn);
            //cmd.ExecuteNonQuery();
            //conn.Close();

            SqlDataAdapter daAns = new SqlDataAdapter("select * from UserAnswer where Sno=0", conn);
            System.Data.DataTable dtAns = new System.Data.DataTable();
            daAns.Fill(dtAns);
            DataRow drAns = dtAns.NewRow();

            drAns["UserId"] = GlobalVar.GlobalUserId;
            drAns["EmailId"] = GlobalVar.GEmailId;
            drAns["SheetName"] = GlobalVar.strSheetName;
            drAns["QNo"] = GlobalVar.QNo;
            drAns["FK_CaseStudyId"] = GlobalVar.CaseStudyId;
            drAns["Status"] = AnsStatus;
            drAns["Question"] = label1.Text;
            dtAns.Rows.Add(drAns);
            SqlCommandBuilder cb = new SqlCommandBuilder(daAns);
            daAns.Update(dtAns);

        }
        private void btnCheckAnswer_Click(object sender, EventArgs e)
        {
            if (label1.Text.Length > 0)
            {
                if(conn.State==ConnectionState.Open)
                {
                    conn.Close();
                }
                conn.Open();

                //string ssql = "delete from UserAnswer  where UserId="+ GlobalVar.GlobalUserId + " and FK_CaseStudyId='"+ GlobalVar.CaseStudyId + "' and QNo="+GlobalVar.QNo+"";
                //SqlCommand cmd = new SqlCommand(ssql, conn);
                //cmd.ExecuteNonQuery();
                //conn.Close();
                //conn.Open();
                //ssql = "Insert into UserAnswer(UserId,EmailId,SheetName,QNo,FK_CaseStudyId) values("+GlobalVar.GlobalUserId +",'"+strEmailId+"','"+ GlobalVar.strSheetName + "',"+QNo+","+ GlobalVar.CaseStudyId +")";
                //cmd = new SqlCommand(ssql, conn);
                //cmd.ExecuteNonQuery();
                //conn.Close();
                groupBox1.Visible = true;
                Grid1Entry(QNo);
                Grid2Entry(QNo);
                Grid3Entry(QNo);
                groupBox1.Visible = false;
                SqlCommand cmd = new SqlCommand();
                if(conn.State == ConnectionState.Open)
                {
                    conn.Close() ;
                }
                conn.Open();
                string sql ="select * from QuestionDetail where  SheetName='"+ GlobalVar.strSheetName + "' and FK_CaseStudyId="+GlobalVar.CaseStudyId+"";
                cmd = new SqlCommand(sql,conn);
               
                SqlDataReader dr= cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        string x = dr["QNo"].ToString();
                        if (x.Length > 0)
                        {
                            if (QNo == Convert.ToInt32(dr["QNo"].ToString()))
                            {
                                intQNo = QNo;
                            }
                        }
                    }
                }
                conn.Close();
            }
        }
        private void btnClose_Click(object sender, EventArgs e)
        {
            btnCloseExcel_Click(sender, e);

            //string fullName = Assembly.GetEntryAssembly().Location;
            //string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
            //string myExcelPath = System.IO.Path.GetDirectoryName(fullName);

            //fileExcel = myExcelPath + "\\Car_Prices_Poland_Student_Book.xlsx"; 
            //Excel.Application xlApp = new Excel.Application();

            //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileExcel);
            //Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1]; 
            
            //Excel.Range xlRange = xlWorksheet.UsedRange;
            //xlApp.DisplayAlerts = false;
            //xlApp.Workbooks.Close();
            //xlApp.Quit();
           

            //        Marshal.ReleaseComObject(xlWorksheet);
            //        Marshal.ReleaseComObject(xlWorkbook);
            //        Marshal.ReleaseComObject(xlApp);


                    this.Close();
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            MsgBox msg = new MsgBox();
            msg.ShowDialog();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {
           
        }

        private void button11_Click_1(object sender, EventArgs e)
        {
            WebBroswer webBroswer = new WebBroswer();
            webBroswer.ShowDialog();

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void lblArrow_Click(object sender, EventArgs e)
        {
            // Assume you have a button (lblArrow) and other controls on the form.

            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int newXPosition;

            if (lblArrow.Text == ">")
            {
                // Move the form to the right side and hide all controls except the button
                newXPosition = screenWidth-80 ;  // Move the form to the right
                lblArrow.Text = "<";

                // Hide all controls except the button (lblArrow in this case)
                foreach (Control control in this.Controls)
                {
                    if (control != lblArrow)  // Exclude the button (lblArrow)
                    {
                        control.Visible = false;
                    }
                }
                //lblArrow.Text = "<";
            }

            else
            {
                // Move the form to the left and show all controls
                newXPosition = 1000;  // Move the form to the left (for example)
                lblArrow.Text = ">";

                // Show all controls
                foreach (Control control in this.Controls)
                {
                    control.Visible = true;
                }

            }
        

            this.SetDesktopLocation(newXPosition, 0);  // Move form to the new X position




        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if (label1.Text.Length > 0)
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
                conn.Open();

                //string ssql = "delete from UserAnswer  where UserId="+ GlobalVar.GlobalUserId + " and FK_CaseStudyId='"+ GlobalVar.CaseStudyId + "' and QNo="+GlobalVar.QNo+"";
                //SqlCommand cmd = new SqlCommand(ssql, conn);
                //cmd.ExecuteNonQuery();
                //conn.Close();
                //conn.Open();
                //ssql = "Insert into UserAnswer(UserId,EmailId,SheetName,QNo,FK_CaseStudyId) values("+GlobalVar.GlobalUserId +",'"+strEmailId+"','"+ GlobalVar.strSheetName + "',"+QNo+","+ GlobalVar.CaseStudyId +")";
                //cmd = new SqlCommand(ssql, conn);
                //cmd.ExecuteNonQuery();
                //conn.Close();
                groupBox1.Visible = true;
                Grid1Entry(QNo);
                Grid2Entry(QNo);
                Grid3Entry(QNo);
                groupBox1.Visible = false;
                SqlCommand cmd = new SqlCommand();
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
                conn.Open();
                string sql = "select * from QuestionDetail where  SheetName='" + GlobalVar.strSheetName + "' and FK_CaseStudyId=" + GlobalVar.CaseStudyId + "";
                cmd = new SqlCommand(sql, conn);

                SqlDataReader dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        string x = dr["QNo"].ToString();
                        if (x.Length > 0)
                        {
                            if (QNo == Convert.ToInt32(dr["QNo"].ToString()))
                            {
                                intQNo = QNo;
                            }
                        }
                    }
                }
                conn.Close();
            }

        }

        //private void pictureBox3_Click(object sender, EventArgs e)
        //{
        //    if (GlobalVar.intOpen == 1)
        //    {
        //        conn.Open();
        //        string selectquery = "SELECT * FROM QuestionDetail where QNo = " + GlobalVar.QNo + " and FK_CaseStudyId=" + GlobalVar.CaseStudyId + "";
        //        SqlCommand cmd = new SqlCommand(selectquery, conn);
        //        SqlDataReader reader1;
        //        reader1 = cmd.ExecuteReader();

        //        if (reader1.Read())
        //        {

        //            //    Form3 frm3 = new Form3();

        //            label3.Visible = false;
        //            label3.Text = reader1.GetValue(6).ToString();
        //            GlobalVar.QLink = reader1.GetValue(12).ToString();
        //            GlobalVar.Qhint = reader1.GetValue(7).ToString();
        //            GlobalVar.frm3 = new Form3();
        //            GlobalVar.frm3.Show();//   frm3.Show();
        //        }
        //        else
        //        {
        //            MessageBox.Show("NO DATA FOUND");
        //        }
        //        conn.Close();
        //        GlobalVar.intOpen = GlobalVar.intOpen + 1;
        //    }
        //}

        private void pictureBox3_Click_1(object sender, EventArgs e)
        {
            if (GlobalVar.intOpen == 1)
            {
                conn.Open();
                string selectquery = "SELECT * FROM QuestionDetail where QNo = " + GlobalVar.QNo + " and FK_CaseStudyId=" + GlobalVar.CaseStudyId + "";
                SqlCommand cmd = new SqlCommand(selectquery, conn);
                SqlDataReader reader1;
                reader1 = cmd.ExecuteReader();

                if (reader1.Read())
                {

                    //    Form3 frm3 = new Form3();

                    label3.Visible = false;
                    label3.Text = reader1.GetValue(6).ToString();
                    GlobalVar.QLink = reader1.GetValue(12).ToString();
                    GlobalVar.Qhint = reader1.GetValue(7).ToString();
                    GlobalVar.frm3 = new Form3();
                    GlobalVar.frm3.Show();//   frm3.Show();
                }
                else
                {
                    MessageBox.Show("NO DATA FOUND");
                }
                conn.Close();
                GlobalVar.intOpen = GlobalVar.intOpen + 1;
            }
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void lblTimer_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        //private void timer2_Tick(object sender, EventArgs e)
        //{



        //        progressBar1.Increment(1);
        //        label5.Text = progressBar1.Value.ToString() + "%";

        //}
    }
}
