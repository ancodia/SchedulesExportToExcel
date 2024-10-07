using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ScheduleExporter
{
    public partial class ExportForm : Form
    {
        private List<string> _originalScheduleNames;
        private List<bool> _selectionStates;

        public string SelectedFilePath { get; private set; }
        public bool WriteAsString { get; private set; }
        public List<string> SelectedSchedules { get; private set; }

        public ExportForm(List<string> scheduleNames)
        {
            InitializeComponent();

            // Default excel filename
            textBoxFilePath.Text = $@"C:\ScheduleExports\{DateTime.UtcNow:yyyyMMddHHmmss}.xlsx";

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
            this.textBoxFilePath = new System.Windows.Forms.TextBox();
            this.labelSchedules = new System.Windows.Forms.Label();
            this.checkedListBoxSchedules = new System.Windows.Forms.CheckedListBox();
            this.checkBoxWriteAsString = new System.Windows.Forms.CheckBox();
            this.labelWriteAsString = new System.Windows.Forms.Label();
            this.buttonWrite = new System.Windows.Forms.Button();
            this.checkBoxSelectAll = new System.Windows.Forms.CheckBox();
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
            this.buttonSelectFile.Location = new System.Drawing.Point(9, 86);
            this.buttonSelectFile.Name = "buttonSelectFile";
            this.buttonSelectFile.Size = new System.Drawing.Size(188, 31);
            this.buttonSelectFile.TabIndex = 2;
            this.buttonSelectFile.Text = "Select Excel File";
            this.buttonSelectFile.UseVisualStyleBackColor = true;
            this.buttonSelectFile.Click += new System.EventHandler(this.buttonSelectFile_Click);
            // 
            // textBoxFilePath
            // 
            this.textBoxFilePath.Location = new System.Drawing.Point(216, 86);
            this.textBoxFilePath.Name = "textBoxFilePath";
            this.textBoxFilePath.Size = new System.Drawing.Size(411, 31);
            this.textBoxFilePath.TabIndex = 3;
            // 
            // labelSchedules
            // 
            this.labelSchedules.AutoSize = true;
            this.labelSchedules.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelSchedules.Location = new System.Drawing.Point(8, 127);
            this.labelSchedules.Name = "labelSchedules";
            this.labelSchedules.Size = new System.Drawing.Size(206, 25);
            this.labelSchedules.TabIndex = 4;
            this.labelSchedules.Text = "Schedules to write";
            // 
            // checkedListBoxSchedules
            // 
            this.checkedListBoxSchedules.CheckOnClick = true;
            this.checkedListBoxSchedules.FormattingEnabled = true;
            this.checkedListBoxSchedules.Location = new System.Drawing.Point(12, 165);
            this.checkedListBoxSchedules.Name = "checkedListBoxSchedules";
            this.checkedListBoxSchedules.ScrollAlwaysVisible = true;
            this.checkedListBoxSchedules.Size = new System.Drawing.Size(615, 228);
            this.checkedListBoxSchedules.TabIndex = 5;
            // 
            // checkBoxWriteAsString
            // 
            this.checkBoxWriteAsString.AutoSize = true;
            this.checkBoxWriteAsString.Checked = true;
            this.checkBoxWriteAsString.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxWriteAsString.Location = new System.Drawing.Point(292, 399);
            this.checkBoxWriteAsString.Name = "checkBoxWriteAsString";
            this.checkBoxWriteAsString.Size = new System.Drawing.Size(115, 29);
            this.checkBoxWriteAsString.TabIndex = 6;
            this.checkBoxWriteAsString.Text = "Yes/No";
            this.checkBoxWriteAsString.UseVisualStyleBackColor = true;
            // 
            // labelWriteAsString
            // 
            this.labelWriteAsString.AutoSize = true;
            this.labelWriteAsString.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelWriteAsString.Location = new System.Drawing.Point(4, 399);
            this.labelWriteAsString.Name = "labelWriteAsString";
            this.labelWriteAsString.Size = new System.Drawing.Size(275, 25);
            this.labelWriteAsString.TabIndex = 7;
            this.labelWriteAsString.Text = "Write numbers as string?";
            // 
            // buttonWrite
            // 
            this.buttonWrite.Location = new System.Drawing.Point(527, 432);
            this.buttonWrite.Name = "buttonWrite";
            this.buttonWrite.Size = new System.Drawing.Size(100, 30);
            this.buttonWrite.TabIndex = 10;
            this.buttonWrite.Text = "Write";
            this.buttonWrite.UseVisualStyleBackColor = true;
            this.buttonWrite.Click += new System.EventHandler(this.buttonWrite_Click);
            // 
            // checkBoxSelectAll
            // 
            this.checkBoxSelectAll.AutoSize = true;
            this.checkBoxSelectAll.Location = new System.Drawing.Point(493, 130);
            this.checkBoxSelectAll.Name = "checkBoxSelectAll";
            this.checkBoxSelectAll.Size = new System.Drawing.Size(134, 29);
            this.checkBoxSelectAll.TabIndex = 8;
            this.checkBoxSelectAll.Text = "Select All";
            this.checkBoxSelectAll.UseVisualStyleBackColor = true;
            this.checkBoxSelectAll.CheckedChanged += new System.EventHandler(this.checkBoxSelectAll_CheckedChanged);
            // 
            // ExportForm
            // 
            this.ClientSize = new System.Drawing.Size(649, 480);
            this.Controls.Add(this.buttonWrite);
            this.Controls.Add(this.labelWriteAsString);
            this.Controls.Add(this.checkBoxWriteAsString);
            this.Controls.Add(this.checkedListBoxSchedules);
            this.Controls.Add(this.labelSchedules);
            this.Controls.Add(this.textBoxFilePath);
            this.Controls.Add(this.buttonSelectFile);
            this.Controls.Add(this.labelTitle);
            this.Controls.Add(this.checkBoxSelectAll);
            this.Name = "ExportForm";
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

            SelectedSchedules = checkedListBoxSchedules.CheckedItems
                .Cast<string>()
                .ToList();

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

        private Label labelTitle;
        private Button buttonSelectFile;
        private TextBox textBoxFilePath;
        private Label labelSchedules;
        private CheckedListBox checkedListBoxSchedules;
        private CheckBox checkBoxWriteAsString;
        private Label labelWriteAsString;
        private Button buttonWrite;
        private CheckBox checkBoxSelectAll;
    }
}
