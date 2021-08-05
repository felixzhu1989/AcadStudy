using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AcadStudy
{
    public partial class FrmMain : Form
    {
        private AcadHelper Acad = new AcadHelper();
        public FrmMain()
        {
            InitializeComponent();
        }

        private void btnReplace_Click(object sender, EventArgs e)
        {
            Acad.ReplaceText(txtFind1.Text,txtReplace1.Text);
            Acad.Zoom();
        }
    }
}
