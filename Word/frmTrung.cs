using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using Image = System.Drawing.Image;

namespace Word
{
    public partial class frmTrung : Form
    {
        float fontSize;
        string fontName;
        float[] size = {15,17,19,20,22,25,28,30,32,35,38,40,42,45,48,50};
        public static RichTextBox rtb;
        string path = "";
        int lengthText = 0;
        bool saved = false;
        Image[] cards;
/*        private System.Drawing.Printing.PrintDocument docToPrint =
    new System.Drawing.Printing.PrintDocument();*/
        
        private PrintDocument printDocument1;
        private string stringToPrint;
        public frmTrung()
        {
            InitializeComponent();
            ReadFontAndSize();
            loadIcons();
            cbbSize.SelectedIndexChanged += new System.EventHandler(this.changeFontAndSize);
            cbbFont.SelectedIndexChanged += new System.EventHandler(this.changeFontAndSize);
            rtb = richTextBox1;

        }
        private void ReadFontAndSize()
        {
            foreach (FontFamily f in FontFamily.Families)
            {
                 cbbFont.Items.Add(f.Name);
            }
            foreach (float f in size)
            {
                cbbSize.Items.Add(f.ToString());
            }
            cbbFont.SelectedIndex = 0;
            cbbSize.SelectedIndex = 0;
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (saved)
            {
                richTextBox1.SaveFile(path);
            }
            else
            {
                try
                {
                    SaveFileDialog saveDialog = new SaveFileDialog();
                    saveDialog.ShowDialog();
                    richTextBox1.SaveFile(saveDialog.FileName);
                 
                    path= saveDialog.FileName;  
                }
                catch (Exception)
                {
                    return;
                }
            }
            saved = true;
        }

        private void changeFontAndSize(object sender, EventArgs e)
        {
            try
            {
                fontName = cbbFont.Text;
                fontSize = float.Parse(cbbSize.Text);
                /*fontName = cbbFont.SelectedItem.ToString();
                fontSize = float.Parse(cbbSize.SelectedItem.ToString());*/
                richTextBox1.SelectionFont = new Font(fontName, fontSize);
            }
            catch (Exception)
            {
                return ;
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            try
            {
                if (richTextBox1.SelectionFont.Style == FontStyle.Bold)
                {
                    richTextBox1.SelectionFont = new Font(fontName,fontSize, FontStyle.Regular);
                    toolStripButton2.Paint -= borderBtn;
                }
                else
                {
                    richTextBox1.SelectionFont = new Font(fontName, fontSize, FontStyle.Bold);
                    toolStripButton2.Paint += borderBtn;
                }
            }
            catch (Exception)
            {
                /*richTextBox1.SelectAll();
                richTextBox1.SelectionFont = new Font(richTextBox1.Font, FontStyle.Regular);
                richTextBox1.DeselectAll();
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.Focus();
                toolStripButton2.Paint -= borderBtn;*/
            }
            
        }
        private void borderBtn(object sender,PaintEventArgs e)
        {
            ToolStripButton btn = (ToolStripButton)sender;
            ControlPaint.DrawBorder(
                   e.Graphics,
                   new Rectangle(0, 0, btn.Width, btn.Height),
                   // or as @LarsTech commented, this works fine too!
                   //  btn.ContentRectangle,
                   Color.Red,
                   ButtonBorderStyle.Solid);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionFont.Style == FontStyle.Italic)
            {
                richTextBox1.SelectionFont = new Font(fontName, fontSize, FontStyle.Regular);
                toolStripButton3.Paint -= borderBtn;
            }
            else
            {
                richTextBox1.SelectionFont = new Font(fontName, fontSize, FontStyle.Italic);
                toolStripButton3.Paint += borderBtn;
            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectionFont.Style == FontStyle.Underline)
            {
                richTextBox1.SelectionFont = new Font(fontName, fontSize, FontStyle.Regular);
                toolStripButton4.Paint -= borderBtn;
            }
            else
            {
                richTextBox1.SelectionFont = new Font(fontName, fontSize, FontStyle.Underline);
                toolStripButton4.Paint += borderBtn;
            }
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.ShowDialog();
            richTextBox1.SelectionColor = colorDialog.Color;
            toolStripButton5.BackColor = colorDialog.Color;
            if(toolStripButton5.BackColor == Color.White)
            {
                toolStripButton5.ForeColor = Color.Black;
                toolStripButton5.Paint += borderBtn;
            }
            else
            {
                toolStripButton5.ForeColor = Color.White;
            }
            
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            lengthText = rtb.Text.Length;
            frmTrung1 frm2 = new frmTrung1();
            frm2.Show();
            /*findText(richTextBox1,"trung");*/
        }
        public static void findText(RichTextBox rtb, string text)
        {
            if (text == "")
            {
                return;
            }
            string[] texts = text.Split(',');
            foreach (string txt in texts)
            {
                int startindex = 0;
                while (startindex < rtb.TextLength)
                {
                    int wordstartIndex = rtb.Find(txt, startindex, RichTextBoxFinds.MatchCase);
                    if (wordstartIndex != -1)
                    {
                        rtb.SelectionStart = wordstartIndex;
                        rtb.SelectionLength = txt.Length;
                        rtb.SelectionBackColor = Color.Yellow;
                    }
                    else
                        break;
                    startindex += wordstartIndex + txt.Length;
                }
            }
            
        }
        public static void replaceText(RichTextBox rtb, string text, string reText)
        {
            string[] texts = text.Split(',');
            foreach (string txt in texts)
            {
                int startindex = 0;
                while (startindex < rtb.TextLength)
                {
                    int wordstartIndex = rtb.Find(txt, startindex, RichTextBoxFinds.None);
                    if (wordstartIndex != -1)
                    {
                        rtb.SelectionStart = wordstartIndex;
                        rtb.SelectionLength = txt.Length;
                        rtb.SelectionBackColor = Color.Yellow;
                        rtb.SelectedText= reText;

                    }
                    else
                        break;
                    /*startindex += wordstartIndex + txt.Length;*/
                }
            }

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            if(lengthText< richTextBox1.Text.Length)
            {

                /*string newText = richTextBox1.Text;
                lengthText = richTextBox1.Text.Length;
                richTextBox1.SelectAll();
                richTextBox1.SelectionBackColor = Color.Transparent;
               
                */
                /*richTextBox1.SelectAll();
                richTextBox1.SelectionBackColor = Color.Transparent;
                richTextBox1.DeselectAll();*/
                /*richTextBox1.Clear();
                richTextBox1.SelectionBackColor = Color.Transparent;
                richTextBox1.Text = newText;
                richTextBox1.SelectionStart = lengthText;
                richTextBox1.Focus();*/
            }

        }
        

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(richTextBox1.SelectedText.Length > 0)
            {
                richTextBox1.Copy();
            }
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Paste();
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectedText.Length > 0)
            {
                richTextBox1.Cut();
            }
        }
        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            richTextBox1.Redo();
        }

        private void toolStripButton7_Click_1(object sender, EventArgs e)
        {
            richTextBox1.Undo();
        }

        private void toolStripButton9_Click(object sender, EventArgs e)
        {
            richTextBox1.ZoomFactor+=1;
        }

        private void toolStripButton10_Click(object sender, EventArgs e)
        {
            if (richTextBox1.ZoomFactor == 1)
            {
                return;
            }
            else
            {
                richTextBox1.ZoomFactor -= 1;
            }
            
        }

        private void homeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /*ToolStripMenuItem menuStrip = (ToolStripMenuItem)sender;*/
            bgMenu(sender);
            toolStripButton9.Visible = false;
            toolStripButton10.Visible = false;
            toolStripButton11.Visible = false;
            toolStripButton12.Visible = false;
            toolStripButton13.Visible = false;
            toolStripButton14.Visible = false;
            toolStripButton15.Visible = false;
            
        }

        private void pageLayoutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bgMenu(sender);
            toolStripButton9.Visible = true;
            toolStripButton10.Visible = true;
            toolStripButton11.Visible = false;
            toolStripButton12.Visible = false;
            toolStripButton13.Visible = false;
            toolStripButton14.Visible = false;
            toolStripButton15.Visible = false;

        }
        
        private void bgMenu(object sender)
        {
            ToolStripMenuItem menuStrip = (ToolStripMenuItem)sender;
            foreach (ToolStripMenuItem item in menuStrip1.Items)
            {
                item.BackColor = Color.FromArgb(192, 255, 192);
            }
            menuStrip.BackColor= Color.White;
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.ShowDialog();
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    richTextBox1.LoadFile(openFileDialog.FileName);
                    saved = true;
                    path= openFileDialog.FileName;

                }
            }
            catch (Exception)
            {
                return ;
            }
            
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dialog = MessageBox.Show("Xác nhận đóng", "Thông báo",
                MessageBoxButtons.YesNoCancel);
            if (dialog==DialogResult.Yes)
            {
                this.Close();
            }

        }

        private void designToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bgMenu(sender);
            toolStripButton14.Visible = true;
            toolStripButton9.Visible = false;
            toolStripButton10.Visible = false;
            toolStripButton11.Visible = false;
            toolStripButton12.Visible = false;
            toolStripButton13.Visible = false;
            toolStripButton15.Visible = true;

        }

        private void pageLayoutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            bgMenu(sender);
            toolStripButton9.Visible = false;
            toolStripButton10.Visible = false;
            toolStripButton11.Visible = true;
            toolStripButton12.Visible = true;
            toolStripButton13.Visible = true;
            toolStripButton14.Visible = false;
            toolStripButton15.Visible = false;

        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bgMenu(sender);
            frmTrung2 frm3=new frmTrung2();
            frm3.Show();
        }

        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bgMenu(sender);
        }

        private void toolStripButton11_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment=HorizontalAlignment.Left;
        }

        private void toolStripButton12_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Center;
        }

        private void toolStripButton13_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Right;
        }

        private void toolStripButton14_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.ShowDialog();
                string img = openFileDialog.FileName;
                Bitmap myBitmap = new Bitmap(img);
                Clipboard.SetDataObject(myBitmap);
                DataFormats.Format format = DataFormats.GetFormat(DataFormats.Bitmap);
                richTextBox1.Paste(format);
            }
            catch (Exception)
            {
                return;
            }

        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            richTextBox1.Visible = true;
            path = "";
            saved = false;
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            richTextBox1.Visible = false;
        }

        private void toolStripButton15_Click(object sender, EventArgs e)
        {
            
            if (listView1.Visible == false)
            {
                listView1.Visible = true;
            }
            else
            {
                listView1.Visible = false;
            }
        }
        private void loadIcons()
        {
            cards = System.IO.Directory.GetFiles(@"../../Resources/icons/", "*.*").Select(f => Image.FromFile(f)).ToArray();
            foreach (Image img in cards)
            {
                imageList1.Images.Add(img);
            }
            listView1.View = View.LargeIcon;
            imageList1.ImageSize = new Size(32,32);
            listView1.LargeImageList = imageList1;
            for (int j = 0; j < imageList1.Images.Count; j++)
            {
                ListViewItem item = new ListViewItem();
                item.ImageIndex = j;
                this.listView1.Items.Add(item);
            }
        }

        

        private void listView1_Click(object sender, EventArgs e)
        {
            int index = listView1.FocusedItem.Index;
            ListViewItem item = listView1.Items[0];
            Image img = item.ImageList.Images[index];
            Bitmap myBitmap = new Bitmap(img);
            Clipboard.SetDataObject(myBitmap);
            DataFormats.Format format = DataFormats.GetFormat(DataFormats.Bitmap);
            richTextBox1.Paste(format);
        }
        PrintDialog printDialog1 =new PrintDialog();

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintDialog pd = new PrintDialog();
            if (pd.ShowDialog() == DialogResult.OK)
            {
                PrintDocument printDocument1 = new PrintDocument();
/*                printDocument1.DefaultPageSettings.PaperSize = new PaperSize("Custum", 500, 500);
*/                printDocument1.PrintPage += new PrintPageEventHandler(this.PrintDocument_PrintPage);
                PrintPreviewDialog printPreviewDialog1 = new PrintPreviewDialog();
                printPreviewDialog1.Document = printDocument1;
                DialogResult result = printPreviewDialog1.ShowDialog();
                if (result == DialogResult.OK)
                    printDocument1.Print();
            }
        }
        private void PrintDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawString(richTextBox1.Text, new Font(richTextBox1.Font.ToString(), richTextBox1.Font.Size), System.Drawing.Brushes.Black, 50, 50);
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Undo();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Redo();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            richTextBox1.Paste();
        }

        private void toolStripButton16_Click(object sender, EventArgs e)
        {
            if(richTextBox1.SelectedText.Length>0)
            richTextBox1.Cut();
        }

        private void toolStripButton17_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectedText.Length > 0)
                richTextBox1.Copy();
        }

        private void toolStripButton18_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            colorDialog.ShowDialog();
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.SelectionBackColor = colorDialog.Color;
                toolStripButton18.BackColor = colorDialog.Color;
            }
            
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void savsAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.ShowDialog();
                richTextBox1.SaveFile(saveDialog.FileName);
                saved = true;
                path = saveDialog.FileName;
            }
            catch (Exception)
            {
                return;
            }
        }
    }
}
