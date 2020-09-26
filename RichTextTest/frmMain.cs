using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RichTextTest
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            rtbMain.Text = "Heading 1:\nThe quick brown fox jumps over.\nthe lazy dog. How about that?\nAnd" +
                " what is more, I think.\nI could write some stuff here.\n" + 
                "Heading 1:\nThe quick brown fox jumps over.\nthe lazy sheep.\n";
            //rtbMain.Find("brown");
            //rtbMain.SelectionFont = f;

            System.Text.RegularExpressions.Regex r;
            
            r = new System.Text.RegularExpressions.Regex("Heading 1:\n.*\n");
            string s = rtbMain.Text;
            var matches = r.Matches(s);
            //while (rr.NextMatch() != null)
            Font f1 = new Font(rtbMain.Font.FontFamily, rtbMain.Font.Size + 2, FontStyle.Bold);
            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                rtbMain.SelectionStart = match.Index;
                rtbMain.SelectionLength = match.Length;
                rtbMain.SelectionFont = f1;
            }
            //rtbMain.Text = rtbMain.Text.Replace("Heading 1:\n", "");
            rtbMain.SelectionStart = 0;
            rtbMain.SelectionLength = 0;

            r = new System.Text.RegularExpressions.Regex("Heading 1:\n");
            while (true)
            {
                var match = r.Match(rtbMain.Text);
                if (match.Success)
                {
                    rtbMain.SelectionStart = match.Index;
                    rtbMain.SelectionLength = match.Length;
                    rtbMain.SelectedText = "";
                }
                else
                {
                    break;
                }

            }


        }

        private void rtbMain_TextChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(rtbMain.Rtf);
        }
    }
}
