using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Reflection;



namespace ExcelApp
{
    public partial class frmUserLogin : Form
    {
        string connectionString = ConfigurationManager.ConnectionStrings["LMS_SQL"].ConnectionString;

        public static string strUserName;
        public static string strUserId; 
        public frmUserLogin()
        {
            InitializeComponent();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            string fullName = Assembly.GetEntryAssembly().Location;
            string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);

            //string myConnectionString = "Data Source=DESKTOP-T064OM6;Initial Catalog=LMS;Persist Security Info=True;User ID=sa; Password=sa@2008";
            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();
            string selectquery = "SELECT UserId FROM Login where EmailId = '" + txtLogin.Text + "' and CaseStudyId = " + Convert.ToInt32(myName) + "";
            SqlCommand cmd = new SqlCommand(selectquery, conn);
            SqlDataReader reader1;
            reader1 = cmd.ExecuteReader();
            if (reader1.Read())
            {
                // strUserName = txtPassword.Text;
                // strUserId = txtLogin.Text;


                GlobalVar.GlobalUserId =Convert.ToInt32(reader1[0].ToString());
                GlobalVar.GEmailId = txtLogin.Text;
                GlobalVar.CaseStudyId = Convert.ToInt32(myName);
                Form2 frmSheet = new Form2();
                frmSheet.Show();
                this.Hide();

                
            }
            else
            {
                MessageBox.Show("User Not Found");
            }
            conn.Close();
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

        private void btnLogin_Click_1(object sender, EventArgs e)
        {
            string fullName = Assembly.GetEntryAssembly().Location;
            string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);

            //string myConnectionString = "Data Source=DESKTOP-T064OM6;Initial Catalog=LMS;Persist Security Info=True;User ID=sa; Password=sa@2008";
            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();
            string selectquery = "SELECT UserId FROM intallium_sa.Login where EmailId = '" + txtLogin.Text + "' and CaseStudyId = " + Convert.ToInt32(myName) + "";
            SqlCommand cmd = new SqlCommand(selectquery, conn);
            SqlDataReader reader1;
            reader1 = cmd.ExecuteReader();
            if (reader1.Read())
            {
                // strUserName = txtPassword.Text;
                // strUserId = txtLogin.Text;


                GlobalVar.GlobalUserId = Convert.ToInt32(reader1[0].ToString());
                GlobalVar.GEmailId = txtLogin.Text;
                GlobalVar.CaseStudyId = Convert.ToInt32(myName);
                GlobalVar.intOpen = 1;
                Form2 frmSheet = new Form2();
                frmSheet.Show();
                this.Hide();


            }
            else
            {
                MessageBox.Show("User Not Found");
            }
            conn.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
#endregion