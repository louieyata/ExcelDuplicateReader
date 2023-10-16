using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;




namespace ExcelFileReader
{
    public partial class DuplicateRowsForm : Form
    {

        public DuplicateRowsForm(Form1 parent)
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;

        }

        private void DuplicateRowsForm_Load(object sender, EventArgs e)
        {
        }

        public void SetMessage(string message)
        {
            richTextBox1.Text = message;
            richTextBox1.Font = new Font("Arial", 13); // Set the font and size
            richTextBox1.ForeColor = Color.Green; // Set the text color
            //richTextBox1.BackColor = Color.LightGray; // Set the background color
            //richTextBox1.SelectionBackColor = Color.Yellow; // Set the selection background color
            richTextBox1.SelectionColor = Color.Black; // Set the selection text color
        }

        public void OKButton1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void ExportButton_Click(object sender, EventArgs e)
        {
            // Create a new Excel package
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                // Add a new worksheet to the Excel package
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Exported Data");

                // Get the text from richTextBox1
                string textToExport = richTextBox1.Text;

                // Split the text by newline character '\n'
                string[] lines = textToExport.Split('\n');

                // Write each line to the Excel worksheet in separate cells
                int row = 1;
                foreach (string line in lines)
                {
                    string[] cells = line.Split('\t'); // Split each line by tab character '\t'

                    for (int col = 1; col <= cells.Length; col++)
                    {
                        worksheet.Cells[row, col].Value = cells[col - 1];
                    }

                    row++;
                }

                // Save the Excel package to a file
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Files|*.xlsx";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                    excelPackage.SaveAs(excelFile);
                    MessageBox.Show("Exported to Excel successfully!");
                }
            }
        }
    }
}
