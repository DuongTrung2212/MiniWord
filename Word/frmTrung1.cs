using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace Word
{
    public partial class frmTrung1 : Form
    {
        public frmTrung1()
        {
            InitializeComponent();
        }

        private void btnFindText_Click(object sender, EventArgs e)
        {
            string text = tbFind.Text;
            frmTrung.findText(frmTrung.rtb, text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string text = tbFindReplay.Text;
            frmTrung.findText(frmTrung.rtb, text);
        }

        private void btnReplay_Click(object sender, EventArgs e)
        {
            string text = tbFindReplay.Text;
            string reText = tbReplay.Text;
            frmTrung.replaceText(frmTrung.rtb, text, reText);
            /*Form1.rtb.SelectedText = "A";*/
        }

        private void btnPageFind_Click(object sender, EventArgs e)
        {
            
            panel1.Visible =true;
            panel2.Visible = false;
            
        }

        private void btnPageReplay_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmTrung1_Load(object sender, EventArgs e)
        {

        }

        private void frmTrung1_FormClosed(object sender, FormClosedEventArgs e)
        {
            
            int lengthText = frmTrung.rtb.Text.Length;
            frmTrung.rtb.SelectAll();
            frmTrung.rtb.SelectionBackColor = Color.Transparent;
            frmTrung.rtb.DeselectAll();
            frmTrung.rtb.SelectionStart = lengthText;
            frmTrung.rtb.Focus();
        }
    }
}
