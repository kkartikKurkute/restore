using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Class1;

namespace ExcelApp
{
    public partial class MsgBox : Form
    {

        public MsgBox()
        {
            InitializeComponent();
        }
        public static string Qhint;
        public static string QLink;

        protected void RePaint()
        {
            GraphicsPath graphicpath = new GraphicsPath();
            graphicpath.StartFigure();
            graphicpath.AddArc(0, 0, 25, 25, 180, 90);
            graphicpath.AddLine(25, 0, this.Width - 25, 0);
            graphicpath.AddArc(this.Width - 25, 0, 25, 25, 270, 90);
            graphicpath.AddLine(this.Width, 25, this.Width, this.Height - 25);
            graphicpath.AddArc(this.Width - 25, this.Height - 25, 25, 25, 0, 90);
            graphicpath.AddLine(this.Width - 25, this.Height, 25, this.Height);
            graphicpath.AddArc(0, this.Height - 25, 25, 25, 90, 90);
            graphicpath.CloseFigure();
            this.Region = new Region(graphicpath);
}
        private void button1_Click(object sender, EventArgs e)
        {

            this.Close();
        }

        private void MsgBox_Load(object sender, EventArgs e)
        {
          
            string fullName = Assembly.GetEntryAssembly().Location;
            string myName = Path.GetDirectoryName("Images");// bGetFileNameWithoutExtension(fullName);
            string myExcelPath = Path.GetDirectoryName(fullName);

            if (GlobalVariables.ModeType == 1)
            {
           

                string filePath = @"\Images\Right.png";
                //picBox.ImageLocation = @"F:\Devendra\Dev\ExcelLiveProject\DesktopAppFormula\ExcelApp\ExcelApp\Images\Right.png";
                picBox.ImageLocation = @""+myExcelPath+"\\Right.png";

                lblMsg.Text = "The answer is Correct...";
                lblMsg.ForeColor = Color.Green;
            }
            if (GlobalVariables.ModeType == 2)
            {
                //picBox.ImageLocation = @"F:\Devendra\Dev\ExcelLiveProject\DesktopAppFormula\ExcelApp\ExcelApp\Images\Incorrect.png";
                picBox.ImageLocation = @"" + myExcelPath + "\\Incorrect.png";
                lblMsg.Text = "The answer is incorrect!";
                lblMsg.ForeColor = Color.Red;
            }
        }
        public class NativeMethods
        {
            [System.Runtime.InteropServices.DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
            public static extern System.IntPtr CreateRoundRectRgn
             (
              int nLeftRect, // x-coordinate of upper-left corner
              int nTopRect, // y-coordinate of upper-left corner
              int nRightRect, // x-coordinate of lower-right corner
              int nBottomRect, // y-coordinate of lower-right corner
              int nWidthEllipse, // height of ellipse
              int nHeightEllipse // width of ellipse
             );

            [System.Runtime.InteropServices.DllImport("gdi32.dll", EntryPoint = "DeleteObject")]
            public static extern bool DeleteObject(System.IntPtr hObject);
            [System.Runtime.InteropServices.DllImportAttribute("user32.dll")]
            public static extern bool ReleaseCapture();

            [System.Runtime.InteropServices.DllImportAttribute("user32.dll")]
            public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        }
        private void btnTryAgain_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        
        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            IntPtr ptr = NativeMethods.CreateRoundRectRgn(15, 15, this.Width, this.Height, 40, 40); // _BoarderRaduis can be adjusted to your needs, try 15 to start.
            this.Region = System.Drawing.Region.FromHrgn(ptr);
            NativeMethods.DeleteObject(ptr);
        }

        #region Make draggable

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                NativeMethods.ReleaseCapture();
                NativeMethods.SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void btnTryAgain_Click_1(object sender, EventArgs e)
        {
            GlobalVar.intClick = 1;
          
            this.Close();
        }

        private void btnSaveNext_Click(object sender, EventArgs e)
        {
           
            GlobalVar.intClick = 2;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (GlobalVar.intOpen == 1)
            {
                GlobalVar.intClick = 2;
                SqlConnection conn = new SqlConnection();
                string connectionString = ConfigurationManager.ConnectionStrings["LMS_SQL"].ConnectionString;
                conn.ConnectionString = connectionString;

                conn.Open();
                string selectquery = "SELECT * FROM QuestionDetail where QNo = " + GlobalVar.QNo + " and FK_CaseStudyId=" + GlobalVar.CaseStudyId + "";
                SqlCommand cmd = new SqlCommand(selectquery, conn);
                SqlDataReader reader1;
                reader1 = cmd.ExecuteReader();

                if (reader1.Read())
                {

                    Form3 frm3 = new Form3();

                    label3.Visible = false;
                    label3.Text = reader1.GetValue(6).ToString();
                    GlobalVar.QLink = reader1.GetValue(12).ToString();
                    GlobalVar.Qhint = reader1.GetValue(7).ToString();
                    GlobalVar.intOpen = GlobalVar.intOpen + 1;
                    GlobalVar.frm3 = new Form3();
                    GlobalVar.frm3.Show();
                }
                else
                {
                    MessageBox.Show("NO DATA FOUND");
                }
                conn.Close();
            }
        }
    }
}
#endregion