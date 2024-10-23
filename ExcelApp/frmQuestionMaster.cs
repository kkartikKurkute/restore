using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ExcelApp
{
    public partial class frmQuestionMaster : Form
    {
        private OleDbConnection connection = new OleDbConnection();
        public frmQuestionMaster()
        {
            InitializeComponent();
            ddlTestName.Items.Add("T1");
            ddlTestName.Items.Add("T2");
            ddlTestName.Items.Add("T3");
            //connection.ConnectionString = "Data Source=DESKTOP-T064OM6;Initial Catalog=LMS;Persist Security Info=True;User ID=sa; Password=sa@2008";
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            try
            {
                string myConnectionString = "Data Source=DESKTOP-T064OM6;Initial Catalog=LMS;Persist Security Info=True;User ID=sa; Password=sa@2008";
                SqlConnection conn = new SqlConnection(myConnectionString);
                conn.Open();

                string Qinsert = "Insert into QuestionMaster (Question,Hint,Link,TestName) values ('" + txtQuestion.Text + "','" + txtHint.Text + "','" + txtLink.Text + "','" + ddlTestName.SelectedItem + "')";
                SqlCommand cmd = new SqlCommand(Qinsert,conn);
                cmd.ExecuteNonQuery();


                
                MessageBox.Show("Data Saved");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error  " + ex);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 frmForm2 = new Form2();
            frmForm2.Show();
        }
    }
}
