using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelApp
{
    public partial class WebBroswer : Form
    {
        public WebBroswer()
        {
            InitializeComponent();
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
        private void WebBroswer_Load(object sender, EventArgs e)
        {
            //string html = System.IO.File.ReadAllText("F:\\Devendra\\Dev\\ExcelLiveProject\\DesktopAppFormula\\ExcelApp\\ExcelApp\\HTMLPage1.html");
            ReplaceInFile("F:\\Devendra\\Dev\\ExcelLiveProject\\DesktopAppFormula\\ExcelApp\\ExcelApp\\HTMLPage1.html", "F:\\Devendra\\Dev\\ExcelLiveProject\\DesktopAppFormula\\ExcelApp\\ExcelApp\\HTMLPage2.html", "opsinhfdsfd sfsfs a", "opsinha");
            //html = html.Replace("opsinha", "OPS SINGH");
            webBrowser1.Navigate("F:\\Devendra\\Dev\\ExcelLiveProject\\DesktopAppFormula\\ExcelApp\\ExcelApp\\HTMLPage2.html");
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }
    }
}
