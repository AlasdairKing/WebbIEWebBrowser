using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WebbIE_I18N
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void btnXMLToExcel_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.FileName = "";
            ofd.Filter = "XML Language Documents|*.Language.xml";
            ofd.CheckFileExists = true;;
            if (ofd.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
            {
                // Name of output
                string name = System.IO.Path.GetFileNameWithoutExtension(ofd.FileName).Replace(".Language", "");

                // Create Excel document
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wb = excelApp.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Worksheet s = wb.Sheets.Add();
                // Add the heading
                s.Cells["1", "A"].Value = "Translation file for " + name;
                s.Cells["1", "A"].Font.Bold = true;;
                s.Cells["2", "A"].Value = "Language";
                s.Cells["2", "B"].Value = this.cboLanguage.Text;
                s.Cells["3", "A"].Value = "Key";
                s.Cells["3", "A"].Font.Bold = true;
                s.Cells["3", "B"].Value = "English value";
                s.Cells["3", "B"].Font.Bold = true;
                s.Cells["3", "C"].Value = "Explanation";
                s.Cells["3", "C"].Font.Bold = true;
                s.Cells["3", "D"].Value = "Translation";
                s.Cells["3", "D"].Font.Bold = true;
                
                // Now load the language XML
                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                try
                {
                    doc.Load(ofd.FileName);
                }
                catch
                {
                    MessageBox.Show("Error loading language XML");
                    return;
                }
                 
                // Go through the language XML file identifying items that do not have a translation
                int rowIndex = 4;
                foreach (System.Xml.XmlNode nodeIterator in doc.DocumentElement.SelectNodes("item"))
                {
                    System.Xml.XmlNodeList languageNodes = nodeIterator.SelectNodes("content[@language=\"" + this.cboLanguage.Text + "\"]");
                    if (languageNodes.Count == 0)
                    {
                        // This is missing a translation into my desired language. Write it to Excel.
                        s.Cells[rowIndex, "A"].Value = nodeIterator.SelectSingleNode("key").InnerText;
                        if (nodeIterator.SelectSingleNode("content[@language='en-gb']") == null)
                        {
                            s.Cells[rowIndex, "B"].Value = nodeIterator.SelectSingleNode("content[@language='en']").InnerText;
                        }
                        else
                        {
                            s.Cells[rowIndex, "B"].Value = nodeIterator.SelectSingleNode("content[@language='en-gb']").InnerText;
                        }
                        if (nodeIterator.SelectSingleNode("explanation") != null)
                        {
                            s.Cells[rowIndex, "C"].Value = nodeIterator.SelectSingleNode("explanation").InnerText;
                        }
                        rowIndex++; // On to the next row.
                    }
                }
                // Save the Excel spreadsheet.
                wb.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\" + name + "-" + this.cboLanguage.Text + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8); // XlFileFormat.xlExcel8);
                Application.DoEvents();
                wb.Close();
                Application.DoEvents();
                excelApp.Quit();
                Application.DoEvents();
                MessageBox.Show(this, "Done!");
                
//    MsgBox "Saved to " & Path
            }
        }
    }
}
