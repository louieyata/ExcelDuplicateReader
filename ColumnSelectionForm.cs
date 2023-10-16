using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelFileReader
{
    public partial class ColumnSelectionForm : Form
    {
        private List<string> columnNames;

        public List<string> SelectedColumns { get; private set; }

        public ColumnSelectionForm(List<string> columnNames)
        {
            InitializeComponent();
            this.columnNames = columnNames;
            SelectedColumns = new List<string>();
        }

        private void ColumnSelectionForm_Load(object sender, EventArgs e)
        {
            SelectedColumns.Clear();
            foreach (string columnName in columnNames)
            {
                var checkBox = new CheckBox();
                checkBox.Text = columnName;
                checkBox.AutoSize = true;
                checkBox.CheckedChanged += CheckBox_CheckedChanged;
                flowLayoutPanel1.Controls.Add(checkBox);
            }
        }

        private void CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox checkBox = (CheckBox)sender;
            if (checkBox.Checked)
            {
                SelectedColumns.Add(checkBox.Text);
            }
            else
            {
                SelectedColumns.Remove(checkBox.Text);
            }
        }

        private void okButton_Click_1(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void cancelButton_Click_1(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
