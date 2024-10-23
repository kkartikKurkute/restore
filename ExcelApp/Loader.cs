using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelApp
{
    public partial class Loader : Form
    {
        public Loader()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Loader_Load(object sender, EventArgs e)
        {
            picBox.ImageLocation = @"F:\Devendra\Dev\ExcelLiveProject\DesktopAppFormula\ExcelApp\ExcelApp\Images\loading-load.gif";

        }
    }
}
