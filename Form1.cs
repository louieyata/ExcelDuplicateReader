using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelDataReader;
//using OfficeOpenXml;
//using OfficeOpenXml.Style;
//using Excel = Microsoft.Office.Interop.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
//using System.Globalization;

namespace ExcelFileReader
{
    public partial class Form1 : Form
    {
        private DataTable dataTable; // Add this field to store the loaded data
        private List<string> selectedColumns;
        private List<string> sheetNames; // Store sheet names here
        private string currentlyLoadedSheetName = ""; // Initialize with an empty string
        private string OriginalFilePath = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            sheetCombo.Items.Clear();
            sheetCombo.SelectedItem = "";
            OpenFileDialog fil = new OpenFileDialog();
            // Set the filter for supported files
            fil.Filter = "Supported files|*.xls;*.xlsx;*.xlsb;*.csv|xls|*.xls|xlsx|*.xlsx|xlsb|*.xlsb|csv|*.csv|All|*.*";
            DialogResult result = fil.ShowDialog();
            if (result == DialogResult.OK)
            {
                dataGridView1.DataSource = null;
                dataGridView1.Columns.Clear();
                dataGridView1.Rows.Clear();
                dataGridView1.Refresh();
                string path = fil.FileName.ToString();
                textBox1.Text = path;
                OriginalFilePath = path;

                // Reset the sheetNames list when opening a new file
                sheetNames = null;

                // Get the list of sheet names from the Excel file if it hasn't been retrieved already
                if (sheetNames == null)
                {
                    sheetNames = GetSheetNames(path);
                    // Populate the sheetCombo ComboBox with sheet names
                    sheetCombo.Items.Clear();
                    foreach (string sheetName in sheetNames)
                    {
                        sheetCombo.Items.Add(sheetName);
                    }
                    sheetCombo.SelectedItem = "VaccinationList";
                }

                // Check if "VaccinationList" sheet exists in the sheet names
                if (sheetNames.Contains("VaccinationList"))
                {
                    // Check if the selected sheet name is different from the currently loaded sheet name
                    if (currentlyLoadedSheetName != "VaccinationList")
                    {
                        dataGridView1.DataSource = null;
                        dataGridView1.Columns.Clear();
                        dataGridView1.Rows.Clear();
                        dataGridView1.Refresh();
                    }

                    // Update the currently loaded sheet name
                    currentlyLoadedSheetName = "VaccinationList";

                    // Proceed with reading the contents from the "VaccinationList" sheet
                    ExcelFileReader(path, "VaccinationList");
                }
                else
                {
                    // Prompt the user to select a sheet name from the list
                    string selectedSheetName = PromptForSheetName(sheetNames);

                    if (!string.IsNullOrWhiteSpace(selectedSheetName))
                    {
                        // Check if the selected sheet name is different from the currently loaded sheet name
                        if (currentlyLoadedSheetName != selectedSheetName)
                        {
                            dataGridView1.DataSource = null;
                            dataGridView1.Columns.Clear();
                            dataGridView1.Rows.Clear();
                            dataGridView1.Refresh();
                        }

                        // Update the currently loaded sheet name
                        currentlyLoadedSheetName = selectedSheetName;

                        ExcelFileReader(path, selectedSheetName);
                    }
                }
            }
            else
            {
                // Handle the case where the user cancels the file selection
                // You can display a message or perform other actions as needed.
            }
        }

        private string PromptForSheetName(List<string> sheetNames)
        {
            string selectedSheetName = null;
            using (var inputForm = new Form())
            using (var tableLayoutPanel = new TableLayoutPanel())
            using (var instructionLabel = new Label()) // Add a label for instructions
            using (var sheetNameListBox = new ListBox())
            using (var okButton = new Button())
            {
                inputForm.Text = "Select Sheet";
                inputForm.Size = new System.Drawing.Size(500, 350);
                inputForm.MinimumSize = inputForm.Size;

                // Center the form on the screen
                inputForm.StartPosition = FormStartPosition.CenterScreen;

                // Create a TableLayoutPanel to manage the layout
                tableLayoutPanel.Dock = DockStyle.Fill;
                tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // Fixed height for the instruction label
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // Fixed height for the button

                // Create the instruction label
                instructionLabel.Dock = DockStyle.Fill;
                instructionLabel.Text = "Please select a sheet name from the list below:";
                //instructionLabel.TextAlign = ContentAlignment.MiddleCenter; // Use TextAlign to set text alignment
                instructionLabel.Font = new System.Drawing.Font("Arial", 12);
                tableLayoutPanel.Controls.Add(instructionLabel, 0, 0);

                // Add controls to the TableLayoutPanel
                sheetNameListBox.Dock = DockStyle.Fill;
                sheetNameListBox.Font = new System.Drawing.Font("Arial", 12);
                sheetNameListBox.BackColor = System.Drawing.Color.White;
                sheetNameListBox.ForeColor = System.Drawing.Color.Black;
                tableLayoutPanel.Controls.Add(sheetNameListBox, 0, 1);

                okButton.Text = "OK";
                okButton.DialogResult = DialogResult.OK;
                okButton.Anchor = AnchorStyles.None; // Center the button
                okButton.Size = new System.Drawing.Size(80, 40); // Set the size here
                tableLayoutPanel.Controls.Add(okButton, 0, 2);

                inputForm.Controls.Add(tableLayoutPanel);

                sheetNameListBox.Items.AddRange(sheetNames.ToArray());

                if (inputForm.ShowDialog() == DialogResult.OK)
                {
                    selectedSheetName = sheetNameListBox.SelectedItem?.ToString();
                }
            }
            return selectedSheetName;
        }

        private List<string> GetSheetNames(string filePath)
        {
            List<string> names = new List<string>();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                var reader = ExcelReaderFactory.CreateReader(stream);

                do
                {
                    names.Add(reader.Name);
                } while (reader.NextResult());
            }

            return names;
        }

        public void ExcelFileReader(string path, string sheetName)
        {
            while (true)
            {
                try
                {
                    dataGridView1.Rows.Clear();
                    using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Find the specified sheet
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration
                            {
                                UseHeaderRow = true
                            }
                        });

                        DataTable table = result.Tables.Cast<DataTable>().FirstOrDefault(t => t.TableName == sheetName);

                        if (table != null)
                        {

                            var dataRows = table.AsEnumerable().Skip(2); //skip the first two rows (based on MR-SIA Template)

                            // Check if the required columns are present
                            var requiredColumns = new List<string> { "Lastname", "Firstname", "Suffixname", "DateofBirth", "Sex", "ActionDate" };
                            var missingColumns = requiredColumns.Except(table.Columns.Cast<DataColumn>().Select(col => col.ColumnName));

                            if (missingColumns.Any())
                            {
                                // Handle missing columns

                                // Columns are present, proceed with displaying data
                                dataGridView1.Columns.AddRange(table.Columns.Cast<DataColumn>()
                                    .Select(col => new DataGridViewTextBoxColumn
                                    {
                                        Name = col.ColumnName,
                                        HeaderText = col.ColumnName
                                    }).ToArray());

                                dataGridView1.Columns["DateofBirth"].DefaultCellStyle.Format = "MM/dd/yyyy";
                                dataGridView1.Columns["ActionDate"].DefaultCellStyle.Format = "MM/dd/yyyy";

                                foreach (DataRow row in dataRows)
                                {
                                    dataGridView1.Rows.Add(row.ItemArray);
                                }

                                // Store the loaded DataTable for later use
                                dataTable = table;
                                break;
                            }
                            else
                            {
                                // Columns are present, proceed with displaying data

                                dataGridView1.Columns.AddRange(table.Columns.Cast<DataColumn>()
                                    .Select(col => new DataGridViewTextBoxColumn
                                    {
                                        Name = col.ColumnName,
                                        HeaderText = col.ColumnName
                                    }).ToArray());

                                dataGridView1.Columns["DateofBirth"].DefaultCellStyle.Format = "MM/dd/yyyy";
                                dataGridView1.Columns["ActionDate"].DefaultCellStyle.Format = "MM/dd/yyyy";

                                foreach (DataRow row in dataRows)
                                {
                                    dataGridView1.Rows.Add(row.ItemArray);
                                }
                                // Store the loaded DataTable for later use
                                dataTable = table;
                                break;
                            }
                        }
                        else
                        {
                            MessageBox.Show($"Sheet '{sheetName}' not found in the Excel file.");

                            // Prompt the user to select another sheet
                            sheetName = PromptForSheetName(sheetNames);

                            if (string.IsNullOrWhiteSpace(sheetName))
                            {
                                // User canceled the sheet selection, exit the loop
                                break;
                            }
                        }

                        // Calculate the width for each column with a minimum width
                        int numberOfColumns = dataGridView1.Columns.Count;
                        int minimumColumnWidth = 100;

                        // Calculate the width, ensuring it's at least the minimum width
                        int columnWidth = Math.Max(dataGridView1.ClientSize.Width / numberOfColumns, minimumColumnWidth);

                        // Set the width for each column
                        foreach (DataGridViewColumn column in dataGridView1.Columns)
                        {
                            column.Width = columnWidth;
                        }
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show($"An error occurred while reading the Excel file. Please choose another sheet.");

                    // Prompt the user to select another sheet
                    sheetName = PromptForSheetName(sheetNames);

                    if (string.IsNullOrWhiteSpace(sheetName))
                    {
                        // User canceled the sheet selection, exit the loop
                        break;
                    }
                }
            }
        }
        private void Button2_Click(object sender, EventArgs e)
        {
            FindDuplicates();
        }


        /*public void FindDuplicates()
        {
            HashSet<string> seen = new HashSet<string>();
            List<string> defaultColumnsToSearch = new List<string> { "Lastname", "Firstname", "Suffixname", "DateofBirth", "Sex", "ActionDate" };

            if (selectedColumns == null)
            {
                MessageBox.Show("Please select the columns for duplicate checking.");
                return;
            }

            // Create a DataTable to hold duplicate rows
            DataTable duplicateRowsDataTable = new DataTable();

            // Create columns in the duplicate rows DataTable based on the selected columns
            foreach (var columnName in selectedColumns)
            {
                duplicateRowsDataTable.Columns.Add(columnName);
            }

            int rowNumber = 2;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                List<string> rowData = new List<string>();

                // Construct the row value string using selected columns
                foreach (var columnName in selectedColumns)
                {
                    var cell = row.Cells[columnName];
                    rowData.Add(cell.Value?.ToString() ?? "");
                }

                string rowValueString = string.Join(", ", rowData);

                if (!string.IsNullOrWhiteSpace(rowValueString))
                {
                    if (seen.Contains(rowValueString))
                    {
                        // Add the duplicate row to the duplicate rows DataTable
                        duplicateRowsDataTable.Rows.Add(rowData.ToArray());
                    }
                    else
                    {
                        seen.Add(rowValueString);
                    }
                }
                rowNumber++;
            }

            if (duplicateRowsDataTable.Rows.Count > 0)
            {
                // Display duplicate rows in the dataGridViewDuplicate in the DuplicateRowsForm
                dataGridViewDuplicate.DataSource = duplicateRowsDataTable;
            }
            else
            {
                MessageBox.Show("No duplicates found.");
            }
        }*/

        public void FindDuplicates()
        {
            HashSet<string> seen = new HashSet<string>();
            StringBuilder duplicateRowsMessage = new StringBuilder();

            List<string> defaultColumnsToSearch = new List<string> { "Lastname", "Firstname", "Suffixname", "DateofBirth", "Sex", "ActionDate" };

            // Check if all default columns are available in the data grid view
            bool allDefaultColumnsExist = defaultColumnsToSearch.All(columnName =>
                dataGridView1.Columns.Cast<DataGridViewColumn>().Any(col => col.HeaderText == columnName));

            if (!allDefaultColumnsExist)
            {
                MessageBox.Show("Default columns (Lastname, Firstname, Suffixname, DateofBirth, Sex, ActionDate) do not exist in the selected sheet. Please select the columns for duplicate checking.");

                // If any of the default columns are missing, show the ColumnSelectionForm
                using (var columnSelectionForm = new ColumnSelectionForm(dataGridView1.Columns.Cast<DataGridViewColumn>().Select(col => col.HeaderText).ToList()))
                {
                    if (columnSelectionForm.ShowDialog() == DialogResult.OK)
                    {
                        selectedColumns = columnSelectionForm.SelectedColumns;
                    }
                }
            }
            else
            {
                // All default columns are available, so use them
                selectedColumns = defaultColumnsToSearch;
            }

            if (selectedColumns != null)
            {
                int rowNumber = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    StringBuilder rowValue = new StringBuilder();

                    // Construct the row value string using selected columns
                    foreach (var columnName in selectedColumns)
                    {
                        var cell = row.Cells[columnName];
                        rowValue.Append(cell.Value?.ToString() ?? "");
                        rowValue.Append(" ");
                    }

                    string rowValueString = rowValue.ToString().Trim();

                    if (!string.IsNullOrWhiteSpace(rowValueString))
                    {
                        if (seen.Contains(rowValueString))
                        {
                            // Append the duplicate row values to the message
                            duplicateRowsMessage.AppendLine($"Duplicate Row {rowNumber}: {rowValueString}");
                        }
                        else
                        {
                            seen.Add(rowValueString);
                        }
                    }
                    rowNumber++;
                }

                if (duplicateRowsMessage.Length > 0)
                {
                    // Display a custom form with the duplicate rows
                    using (DuplicateRowsForm duplicateForm = new DuplicateRowsForm(this))
                    {
                        duplicateForm.SetMessage(duplicateRowsMessage.ToString());
                        duplicateForm.ShowDialog();
                    }
                }
                else
                {
                    MessageBox.Show("No duplicates found.");
                }
                selectedColumns.Clear();
            }

        }

        private void button3_Click(object sender, EventArgs e)//remove duplicates button
        {
            FindDuplicatesAndDelete();
        }

        private void FindDuplicatesAndDelete()
        {
            if (dataTable == null)
            {
                MessageBox.Show("Please load an Excel file first.");
                return;
            }

            // Define the default column names to search for duplicates
            List<string> defaultColumnsToSearch = new List<string> { "Lastname", "Firstname", "Suffixname", "DateofBirth", "Sex", "ActionDate" };

            // Check if all default columns are available in the data table
            bool allDefaultColumnsExist = defaultColumnsToSearch.All(columnName =>
                dataTable.Columns.Cast<DataColumn>().Any(col => col.ColumnName == columnName));

            if (!allDefaultColumnsExist)
            {
                // If any of the default columns are missing, show the ColumnSelectionForm
                using (var columnSelectionForm = new ColumnSelectionForm(dataGridView1.Columns.Cast<DataGridViewColumn>().Select(col => col.HeaderText).ToList()))
                {
                    if (columnSelectionForm.ShowDialog() == DialogResult.OK)
                    {
                        selectedColumns = columnSelectionForm.SelectedColumns;
                    }
                    else
                    {
                        selectedColumns = null;
                        return;
                    }
                }
            }
            else
            {
                // All default columns are available, so use them
                selectedColumns = defaultColumnsToSearch;
            }

            if (selectedColumns != null)
            {
                // Check for duplicates based on the selected columns
                var duplicates = dataTable.AsEnumerable().Skip(2)
                    .GroupBy(row => string.Join(",", selectedColumns.Select(col => row[col])), StringComparer.OrdinalIgnoreCase)
                    .Where(group => group.Count() > 1)
                    .SelectMany(group => group.Skip(1)); // Select all but the first occurrence

                // Check if duplicates were found before saving
                if (duplicates.Any())
                {
                    // Remove duplicate rows

                    foreach (var duplicateRow in duplicates.ToList())
                    {
                        dataTable.Rows.Remove(duplicateRow);
                    }

                    // Refresh DataGridView
                    dataGridView1.DataSource = null;
                    dataGridView1.Columns.Clear();
                    dataGridView1.Rows.Clear();
                    dataGridView1.Refresh();
                    dataGridView1.DataSource = dataTable;

                    dataGridView1.DataSource = dataTable.AsEnumerable().Skip(2).CopyToDataTable();
                    dataGridView1.Columns["DateofBirth"].DefaultCellStyle.Format = "MM/dd/yyyy";
                    dataGridView1.Columns["ActionDate"].DefaultCellStyle.Format = "MM/dd/yyyy";

                    // Save changes back to the Excel file
                    if (SaveChangesToExcelFile())
                    {
                        MessageBox.Show("Duplicate rows removed and changes saved successfully.");
                    }
                    else
                    {
                      MessageBox.Show("Failed to save changes to the Excel file.");
                    }
                }
                else
                {
                    MessageBox.Show("No duplicates found; no changes to save.");
                }
            }
        }

        private bool SaveChangesToExcelFile()
        {
            if (dataTable == null)
            {
                return false;
            }

            try
            {
                string copyFilePath = Path.Combine(Path.GetDirectoryName(OriginalFilePath), "Copy_" + Path.GetFileName(OriginalFilePath));
                var dataTableCopy = dataTable.AsEnumerable().Skip(2).CopyToDataTable(); // Create a copy of the dataTable and skip the first row when saving

                File.Copy(OriginalFilePath, copyFilePath, true); // Copy the original file to a new file

                using (FileStream fs = new FileStream(copyFilePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    IWorkbook workbook = new XSSFWorkbook(fs); // Load the Excel workbook from the copied file
                    ISheet worksheet = workbook.GetSheet("VaccinationList"); // Get the "VaccinationList" worksheet

                    // Clear all data in the "VaccinationList" sheet except the first row
                    for (int i = worksheet.LastRowNum; i >= 3; i--)
                    {
                        IRow row = worksheet.GetRow(i);
                        worksheet.RemoveRow(row);
                    }

                    // Define a cell style for center alignment
                    ICellStyle centerAlignmentStyle = workbook.CreateCellStyle();
                    centerAlignmentStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    centerAlignmentStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

                    for (int i = 0; i < dataTableCopy.Rows.Count; i++)
                    {
                        IRow dataRow = worksheet.CreateRow(i + 3); // Start from index 3 to skip the header row

                        for (int j = 0; j < dataTableCopy.Columns.Count; j++)
                        {
                            ICell cell = dataRow.CreateCell(j);
                            object cellValue = dataTableCopy.Rows[i][j];

                            if (j == 5 || j == 18 || j == 20) // Column indices for date columns
                            {
                                if (cellValue is DateTime dateValue)
                                {
                                    // Format the date as a string
                                    string formattedDate = dateValue.ToString("MM/dd/yyyy");

                                    cell.SetCellValue(formattedDate);

                                    // Set the cell type to STRING to ensure Excel recognizes it as text
                                    cell.SetCellType(CellType.String);
                                }
                                else
                                {
                                    cell.SetCellValue(cellValue.ToString());
                                }
                            }
                            else
                            {
                                cell.SetCellValue(cellValue.ToString());
                            }

                            // Apply the center alignment style to the cell
                            cell.CellStyle = centerAlignmentStyle;
                        }
                    }




                    /*// Create a new row at index 3 and load data from the copied DataTable
                    for (int i = 0; i < dataTableCopy.Rows.Count; i++)
                    {
                        IRow dataRow = worksheet.CreateRow(i + 3); // Start from index 3 to skip the header row

                        for (int j = 0; j < dataTableCopy.Columns.Count; j++)
                        {
                            ICell cell = dataRow.CreateCell(j);
                            object cellValue = dataTableCopy.Rows[i][j];

                            if (j == 5 || j == 18 || j == 20) // Column indices for date columns
                            {
                                if (cellValue is DateTime dateValue)
                                {
                                    // Format the date as a string
                                    string formattedDate = dateValue.ToString("MM/dd/yyyy");

                                    cell.SetCellValue(formattedDate);

                                    // Set the cell type to STRING to ensure Excel recognizes it as text
                                    cell.SetCellType(CellType.String);
                                }
                                else
                                {
                                    cell.SetCellValue(cellValue.ToString());
                                }
                            }
                            else
                            {
                                cell.SetCellValue(cellValue.ToString());
                            }
                        }
                    }*/




                    // Save the changes back to the file
                    using (FileStream fileStream = new FileStream(copyFilePath, FileMode.Create))
                    {
                        workbook.Write(fileStream);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            // Display an error message in a message box
                    MessageBox.Show("An error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Check if a sheet is selected in the combo box
            if (sheetCombo.SelectedItem != null)
            {
                string selectedSheetName = sheetCombo.SelectedItem.ToString();

                // Check if the selected sheet exists in the loaded sheet names
                if (sheetNames.Contains(selectedSheetName))
                {
                    // Clear the previous data from the DataGridView
                    dataGridView1.DataSource = null;
                    dataGridView1.Columns.Clear();
                    dataGridView1.Rows.Clear();
                    dataGridView1.Refresh();

                    // Call ExcelFileReader with the selected sheet name to display data
                    ExcelFileReader(textBox1.Text, selectedSheetName); // Use the selected Excel file path from textBox1
                }
                else
                {
                    // Handle the case where the selected sheet doesn't exist in the loaded sheet names
                    MessageBox.Show($"Sheet '{selectedSheetName}' does not exist in the loaded Excel file.");
                }
            }
        }
    }
}
