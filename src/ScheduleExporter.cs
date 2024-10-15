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
            labelTitle = new Label();
            buttonSelectFile = new Button();
            textBoxFilePath = new TextBox();
            labelSchedules = new Label();
            checkedListBoxSchedules = new CheckedListBox();
            checkBoxWriteAsString = new CheckBox();
            labelWriteAsString = new Label();
            buttonWrite = new Button();
            checkBoxSelectAll = new CheckBox();
            SuspendLayout();
            // 
            // labelTitle
            // 
            labelTitle.AutoSize = true;
            labelTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            labelTitle.Location = new System.Drawing.Point(12, 20);
            labelTitle.Name = "labelTitle";
            labelTitle.Size = new System.Drawing.Size(545, 44);
            labelTitle.TabIndex = 0;
            labelTitle.Text = "Export schedules to Excel file";
            // 
            // buttonSelectFile
            // 
            buttonSelectFile.Location = new System.Drawing.Point(9, 86);
            buttonSelectFile.Name = "buttonSelectFile";
            buttonSelectFile.Size = new System.Drawing.Size(188, 31);
            buttonSelectFile.TabIndex = 2;
            buttonSelectFile.Text = "Select Excel File";
            buttonSelectFile.UseVisualStyleBackColor = true;
            buttonSelectFile.Click += buttonSelectFile_Click;
            // 
            // textBoxFilePath
            // 
            textBoxFilePath.Location = new System.Drawing.Point(216, 86);
            textBoxFilePath.Name = "textBoxFilePath";
            textBoxFilePath.Size = new System.Drawing.Size(411, 39);
            textBoxFilePath.TabIndex = 3;
            // 
            // labelSchedules
            // 
            labelSchedules.AutoSize = true;
            labelSchedules.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            labelSchedules.Location = new System.Drawing.Point(8, 127);
            labelSchedules.Name = "labelSchedules";
            labelSchedules.Size = new System.Drawing.Size(206, 25);
            labelSchedules.TabIndex = 4;
            labelSchedules.Text = "Schedules to write";
            // 
            // checkedListBoxSchedules
            // 
            checkedListBoxSchedules.CheckOnClick = true;
            checkedListBoxSchedules.FormattingEnabled = true;
            checkedListBoxSchedules.Location = new System.Drawing.Point(12, 165);
            checkedListBoxSchedules.Name = "checkedListBoxSchedules";
            checkedListBoxSchedules.ScrollAlwaysVisible = true;
            checkedListBoxSchedules.Size = new System.Drawing.Size(615, 220);
            checkedListBoxSchedules.TabIndex = 5;
            // 
            // checkBoxWriteAsString
            // 
            checkBoxWriteAsString.AutoSize = true;
            checkBoxWriteAsString.Checked = true;
            checkBoxWriteAsString.CheckState = CheckState.Checked;
            checkBoxWriteAsString.Location = new System.Drawing.Point(292, 399);
            checkBoxWriteAsString.Name = "checkBoxWriteAsString";
            checkBoxWriteAsString.Size = new System.Drawing.Size(121, 36);
            checkBoxWriteAsString.TabIndex = 6;
            checkBoxWriteAsString.Text = "Yes/No";
            checkBoxWriteAsString.UseVisualStyleBackColor = true;
            // 
            // labelWriteAsString
            // 
            labelWriteAsString.AutoSize = true;
            labelWriteAsString.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            labelWriteAsString.Location = new System.Drawing.Point(4, 399);
            labelWriteAsString.Name = "labelWriteAsString";
            labelWriteAsString.Size = new System.Drawing.Size(275, 25);
            labelWriteAsString.TabIndex = 7;
            labelWriteAsString.Text = "Write numbers as string?";
            // 
            // buttonWrite
            // 
            buttonWrite.Location = new System.Drawing.Point(527, 432);
            buttonWrite.Name = "buttonWrite";
            buttonWrite.Size = new System.Drawing.Size(100, 30);
            buttonWrite.TabIndex = 10;
            buttonWrite.Text = "Write";
            buttonWrite.UseVisualStyleBackColor = true;
            buttonWrite.Click += buttonWrite_Click;
            // 
            // checkBoxSelectAll
            // 
            checkBoxSelectAll.AutoSize = true;
            checkBoxSelectAll.Location = new System.Drawing.Point(493, 130);
            checkBoxSelectAll.Name = "checkBoxSelectAll";
            checkBoxSelectAll.Size = new System.Drawing.Size(144, 36);
            checkBoxSelectAll.TabIndex = 8;
            checkBoxSelectAll.Text = "Select All";
            checkBoxSelectAll.UseVisualStyleBackColor = true;
            checkBoxSelectAll.CheckedChanged += checkBoxSelectAll_CheckedChanged;
            // 
            // ExportForm
            // 
            ClientSize = new System.Drawing.Size(649, 480);
            Controls.Add(buttonWrite);
            Controls.Add(labelWriteAsString);
            Controls.Add(checkBoxWriteAsString);
            Controls.Add(checkedListBoxSchedules);
            Controls.Add(labelSchedules);
            Controls.Add(textBoxFilePath);
            Controls.Add(buttonSelectFile);
            Controls.Add(labelTitle);
            Controls.Add(checkBoxSelectAll);

            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            Name = "ExportForm";
            Text = "Schedule Exporter";
            ResumeLayout(false);
            PerformLayout();

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
