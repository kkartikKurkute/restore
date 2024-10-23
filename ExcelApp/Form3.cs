using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelApp
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
            this.FormClosing += Form3_Closing;
        }
        private void Form3_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            GlobalVar.intOpen = 1; 
        }
            private void Form3_Load(object sender, EventArgs e)
        {
            txtHintDesc.Text  = GlobalVar.Qhint;
            lbltitle2.Text = GlobalVar.QLink;

            string fullName = Assembly.GetEntryAssembly().Location;
            string myName = System.IO.Path.GetFileNameWithoutExtension(fullName);
            string myExcelPath = System.IO.Path.GetDirectoryName(fullName);


            //string html = System.IO.File.ReadAllText("F:\\Devendra\\Dev\\ExcelLiveProject\\DesktopAppFormula\\ExcelApp\\ExcelApp\\HTMLPage1.html");
            //ReplaceInFile("F:\\Devendra\\Dev\\ExcelLiveProject\\DesktopAppFormula\\ExcelApp\\ExcelApp\\HTMLPage1.html", "F:\\Devendra\\Dev\\ExcelLiveProject\\DesktopAppFormula\\ExcelApp\\ExcelApp\\HTMLPage2.html", txtHintDesc.Text, "opsinha");
            ReplaceInFile(myExcelPath+"\\HTMLPage1.html", myExcelPath + "\\HTMLPage2.html", txtHintDesc.Text, "opsinha");

            //html = html.Replace("opsinha", "OPS SINGH");
            webBrowser1.Navigate(myExcelPath + "\\HTMLPage2.html");

        }
        public void ReplaceInFile(string filePath, string filePath1, string replaceText, string findText)
        {
            try
            {
                // Read the complete file and replace the text
                using (StreamReader reader = new StreamReader(filePath))
                {
                    string content = reader.ReadToEnd();
                    content = Regex.Replace(content, findText, replaceText);

                    // Write the content back to the file
                    using (StreamWriter writer = new StreamWriter(filePath1))
                    {
                        writer.Write(content);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                // Handle the exception as needed (logging, etc.)
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(lbltitle2.Text);
            //string html = "<html><head>";
            //html += "<meta content='IE=Edge' http-equiv='X-UA-Compatible'/>";
            //html += "<iframe id='video' src= 'https://www.youtube.com/embed/{0}' width='420' height='250' frameborder='0' allowfullscreen></iframe>";
            //html += "</body></html>";
            //this.webBrowser1.DocumentText = string.Format(html, txtUrl.Text.Split('=')[1]);
        }

      
    }
}
