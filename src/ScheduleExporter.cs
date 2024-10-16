using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ScheduleExporter
{
    public partial class ExportForm : Form
    {
        private Label labelTitle;
        private Button buttonSelectFile;
        private Button buttonNewFile;
        private TextBox textBoxFilePath;
        private Label labelSchedules;
        private CheckedListBox checkedListBoxSchedules;
        private CheckBox checkBoxWriteAsString;
        private Label labelWriteAsString;
        private Button buttonWrite;
        private CheckBox checkBoxSelectAll;
        private Label lableFilePath;
        private Label labelExcludeEmpty;
        private CheckBox checkBoxExcludeEmpty;

        private List<string> _originalScheduleNames;
        private List<bool> _selectionStates;

        public string SelectedFilePath { get; private set; }
        public bool WriteAsString { get; private set; }
        public bool ExcludeEmptyRows { get; private set; }
        public List<string> SelectedSchedules { get; private set; }

        public ExportForm(List<string> scheduleNames)
        {
            InitializeComponent();

            // Store the original schedule names
            _originalScheduleNames = scheduleNames;

            // Initialize selection states
            _selectionStates = new List<bool>(new bool[scheduleNames.Count]);

            // Populate the CheckedListBox with the schedule names
            PopulateSchedules(scheduleNames);
        }

        // Populates the CheckedListBox with the provided schedule names
        private void PopulateSchedules(List<string> scheduleNames)
        {
            checkedListBoxSchedules.Items.Clear();

            foreach (var schedule in scheduleNames)
            {
                checkedListBoxSchedules.Items.Add(schedule);
            }

            // Restore selection states after repopulating
            for (int i = 0; i < scheduleNames.Count; i++)
            {
                checkedListBoxSchedules.SetItemChecked(i, _selectionStates[i]);
            }
        }

        private void InitializeComponent()
        {
            this.labelTitle = new System.Windows.Forms.Label();
            this.buttonSelectFile = new System.Windows.Forms.Button();
            this.buttonNewFile = new System.Windows.Forms.Button();
            this.textBoxFilePath = new System.Windows.Forms.TextBox();
            this.labelSchedules = new System.Windows.Forms.Label();
            this.checkedListBoxSchedules = new System.Windows.Forms.CheckedListBox();
            this.checkBoxWriteAsString = new System.Windows.Forms.CheckBox();
            this.labelWriteAsString = new System.Windows.Forms.Label();
            this.buttonWrite = new System.Windows.Forms.Button();
            this.checkBoxSelectAll = new System.Windows.Forms.CheckBox();
            this.lableFilePath = new System.Windows.Forms.Label();
            this.labelExcludeEmpty = new System.Windows.Forms.Label();
            this.checkBoxExcludeEmpty = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // labelTitle
            // 
            this.labelTitle.AutoSize = true;
            this.labelTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTitle.Location = new System.Drawing.Point(12, 20);
            this.labelTitle.Name = "labelTitle";
            this.labelTitle.Size = new System.Drawing.Size(545, 44);
            this.labelTitle.TabIndex = 0;
            this.labelTitle.Text = "Export schedules to Excel file";
            // 
            // buttonSelectFile
            // 
            this.buttonSelectFile.Location = new System.Drawing.Point(301, 86);
            this.buttonSelectFile.Name = "buttonSelectFile";
            this.buttonSelectFile.Size = new System.Drawing.Size(283, 40);
            this.buttonSelectFile.TabIndex = 2;
            this.buttonSelectFile.Text = "Select File";
            this.buttonSelectFile.UseVisualStyleBackColor = true;
            this.buttonSelectFile.Click += new System.EventHandler(this.buttonSelectFile_Click);
            // 
            // buttonNewFile
            // 
            this.buttonNewFile.Location = new System.Drawing.Point(12, 86);
            this.buttonNewFile.Name = "buttonNewFile";
            this.buttonNewFile.Size = new System.Drawing.Size(283, 40);
            this.buttonNewFile.TabIndex = 1;
            this.buttonNewFile.Text = "New File";
            this.buttonNewFile.UseVisualStyleBackColor = true;
            this.buttonNewFile.Click += new System.EventHandler(this.buttonNewFile_Click);
            // 
            // textBoxFilePath
            // 
            this.textBoxFilePath.Location = new System.Drawing.Point(129, 148);
            this.textBoxFilePath.Name = "textBoxFilePath";
            this.textBoxFilePath.Size = new System.Drawing.Size(455, 31);
            this.textBoxFilePath.TabIndex = 3;
            // 
            // labelSchedules
            // 
            this.labelSchedules.AutoSize = true;
            this.labelSchedules.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSchedules.Location = new System.Drawing.Point(8, 207);
            this.labelSchedules.Name = "labelSchedules";
            this.labelSchedules.Size = new System.Drawing.Size(223, 32);
            this.labelSchedules.TabIndex = 0;
            this.labelSchedules.Text = "Schedules to write";
            // 
            // checkedListBoxSchedules
            // 
            this.checkedListBoxSchedules.CheckOnClick = true;
            this.checkedListBoxSchedules.FormattingEnabled = true;
            this.checkedListBoxSchedules.Location = new System.Drawing.Point(12, 245);
            this.checkedListBoxSchedules.Name = "checkedListBoxSchedules";
            this.checkedListBoxSchedules.ScrollAlwaysVisible = true;
            this.checkedListBoxSchedules.Size = new System.Drawing.Size(572, 396);
            this.checkedListBoxSchedules.TabIndex = 5;
            // 
            // checkBoxWriteAsString
            // 
            this.checkBoxWriteAsString.AutoSize = true;
            this.checkBoxWriteAsString.Checked = true;
            this.checkBoxWriteAsString.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxWriteAsString.Location = new System.Drawing.Point(315, 660);
            this.checkBoxWriteAsString.Name = "checkBoxWriteAsString";
            this.checkBoxWriteAsString.Size = new System.Drawing.Size(115, 29);
            this.checkBoxWriteAsString.TabIndex = 6;
            this.checkBoxWriteAsString.Text = "Yes/No";
            this.checkBoxWriteAsString.UseVisualStyleBackColor = true;
            // 
            // labelWriteAsString
            // 
            this.labelWriteAsString.AutoSize = true;
            this.labelWriteAsString.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelWriteAsString.Location = new System.Drawing.Point(5, 655);
            this.labelWriteAsString.Name = "labelWriteAsString";
            this.labelWriteAsString.Size = new System.Drawing.Size(301, 32);
            this.labelWriteAsString.TabIndex = 0;
            this.labelWriteAsString.Text = "Write numbers as strings";
            // 
            // buttonWrite
            // 
            this.buttonWrite.Location = new System.Drawing.Point(301, 761);
            this.buttonWrite.Name = "buttonWrite";
            this.buttonWrite.Size = new System.Drawing.Size(283, 47);
            this.buttonWrite.TabIndex = 8;
            this.buttonWrite.Text = "Write";
            this.buttonWrite.UseVisualStyleBackColor = true;
            this.buttonWrite.Click += new System.EventHandler(this.buttonWrite_Click);
            // 
            // checkBoxSelectAll
            // 
            this.checkBoxSelectAll.AutoSize = true;
            this.checkBoxSelectAll.Location = new System.Drawing.Point(440, 203);
            this.checkBoxSelectAll.Name = "checkBoxSelectAll";
            this.checkBoxSelectAll.Size = new System.Drawing.Size(134, 29);
            this.checkBoxSelectAll.TabIndex = 4;
            this.checkBoxSelectAll.Text = "Select All";
            this.checkBoxSelectAll.UseVisualStyleBackColor = true;
            this.checkBoxSelectAll.CheckedChanged += new System.EventHandler(this.checkBoxSelectAll_CheckedChanged);
            // 
            // lableFilePath
            // 
            this.lableFilePath.AutoSize = true;
            this.lableFilePath.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lableFilePath.Location = new System.Drawing.Point(12, 148);
            this.lableFilePath.Name = "lableFilePath";
            this.lableFilePath.Size = new System.Drawing.Size(111, 32);
            this.lableFilePath.TabIndex = 0;
            this.lableFilePath.Text = "File path";
            // 
            // labelExcludeEmpty
            // 
            this.labelExcludeEmpty.AutoSize = true;
            this.labelExcludeEmpty.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelExcludeEmpty.Location = new System.Drawing.Point(6, 711);
            this.labelExcludeEmpty.Name = "labelExcludeEmpty";
            this.labelExcludeEmpty.Size = new System.Drawing.Size(277, 32);
            this.labelExcludeEmpty.TabIndex = 0;
            this.labelExcludeEmpty.Text = "Exclude all empty rows";
            // 
            // checkBoxExcludeEmpty
            // 
            this.checkBoxExcludeEmpty.AutoSize = true;
            this.checkBoxExcludeEmpty.Location = new System.Drawing.Point(315, 716);
            this.checkBoxExcludeEmpty.Name = "checkBoxExcludeEmpty";
            this.checkBoxExcludeEmpty.Size = new System.Drawing.Size(115, 29);
            this.checkBoxExcludeEmpty.TabIndex = 7;
            this.checkBoxExcludeEmpty.Text = "Yes/No";
            this.checkBoxExcludeEmpty.UseVisualStyleBackColor = true;
            // 
            // ExportForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(599, 820);
            this.Controls.Add(this.labelExcludeEmpty);
            this.Controls.Add(this.checkBoxExcludeEmpty);
            this.Controls.Add(this.lableFilePath);
            this.Controls.Add(this.buttonWrite);
            this.Controls.Add(this.labelWriteAsString);
            this.Controls.Add(this.checkBoxWriteAsString);
            this.Controls.Add(this.checkedListBoxSchedules);
            this.Controls.Add(this.labelSchedules);
            this.Controls.Add(this.textBoxFilePath);
            this.Controls.Add(this.buttonSelectFile);
            this.Controls.Add(this.buttonNewFile);
            this.Controls.Add(this.labelTitle);
            this.Controls.Add(this.checkBoxSelectAll);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ExportForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Schedule Exporter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void buttonSelectFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxFilePath.Text = openFileDialog.FileName;
            }
        }

        private void buttonWrite_Click(object sender, EventArgs e)
        {
            // Set the properties before closing the form
            SelectedFilePath = textBoxFilePath.Text;
            WriteAsString = checkBoxWriteAsString.Checked;
            ExcludeEmptyRows = checkBoxExcludeEmpty.Checked;

            SelectedSchedules = checkedListBoxSchedules.CheckedItems
                .Cast<string>()
                .ToList();

            // Check if SelectedFilePath or SelectedSchedules is empty
            if (string.IsNullOrWhiteSpace(SelectedFilePath) || SelectedSchedules.Count == 0)
            {
                MessageBox.Show(
                    "Please create or select a file and at least one schedule.", 
                    "Input Required", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Warning
                );
                return;
            }

            // Check if the file at SelectedFilePath is currently open
            if (Helper.IsFileOpen(textBoxFilePath.Text))
            {
                MessageBox.Show(
                    "The selected file is currently open. Please close it before proceeding.", 
                    "File In Use", 
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void checkBoxSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            bool isChecked = checkBoxSelectAll.Checked;

            for (int i = 0; i < checkedListBoxSchedules.Items.Count; i++)
            {
                checkedListBoxSchedules.SetItemChecked(i, isChecked);
            }
        }

        private void buttonNewFile_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel Files|*.xlsx";
                saveFileDialog.Title = "Save an Excel File";
                saveFileDialog.FileName = "SchedulesExport";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        ExcelPackage.LicenseContext = LicenseContext.Commercial;
                        FileInfo newFile = new FileInfo(saveFileDialog.FileName);
                        using (ExcelPackage excelPackage = new ExcelPackage(newFile))
                        {
                            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                            excelPackage.Save();
                        }
                        textBoxFilePath.Text = saveFileDialog.FileName;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                }
            }
        }
    }
}
