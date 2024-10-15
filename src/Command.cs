using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using ScheduleExporter;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SchedulesExcelExport
{
    [Transaction(TransactionMode.Manual)]
    public partial class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            Settings docSettings = doc.Settings;

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            List<ViewSchedule> schedules = collector
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(schedule => !schedule.Name.Contains("<Revision"))
                .OrderBy(vs => vs.Name)
                .ToList();

            List<string> scheduleNames = schedules
                .Select(schedule => schedule.Name)
                .ToList();

            try
            {
                // Initialize and show the WinForm
                ExportForm exportForm = new(scheduleNames);
                var result = exportForm.ShowDialog();

                if (result == DialogResult.OK)
                {
                    string filePath = exportForm.SelectedFilePath;
                    bool writeAsString = exportForm.WriteAsString;
                    List<string> selectedScheduleNames = exportForm.SelectedSchedules;

                    
                    List<ViewSchedule> selectedSchedules = schedules
                        .Where(schedule => selectedScheduleNames.Contains(schedule.Name))
                        .ToList();

                    ExportToExcel(selectedSchedules, filePath, writeAsString);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        public void ExportToExcel(List<ViewSchedule> schedules, string filePath = null, bool numbersAsStrings = true)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            using (ExcelPackage excelPackage = new())
            {
                FileInfo excelFile = new(filePath);

                if (excelFile.Exists && IsFileOpen(filePath))
                {
                    MessageBox.Show("The Excel file is currently open. Please close it and try again.", "File In Use");
                    return;
                }

                // Load content of existing file if it exists
                if (excelFile.Exists)
                {
                    using (FileStream stream = new(filePath, FileMode.Open))
                    {
                        excelPackage.Load(stream);
                    }
                }

                foreach (ViewSchedule schedule in schedules)
                {
                    string scheduleName = SanitizeWorksheetName(schedule.Name);
                    var worksheet = excelPackage.Workbook.Worksheets[scheduleName] ?? excelPackage.Workbook.Worksheets.Add(scheduleName);

                    var data = PrepareScheduleData(schedule, numbersAsStrings);

                    // Write data directly to the worksheet
                    var startRow = 1;
                    var startCol = 1;
                    worksheet.Cells[startRow, startCol].LoadFromArrays(data.Select(row => row.ToArray()).ToArray());
                }

                Directory.CreateDirectory(Path.GetDirectoryName(filePath));
                excelPackage.SaveAs(excelFile);
            }

            // Open the saved Excel file
            Process.Start(new ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            });
        }

        private List<List<object>> PrepareScheduleData(ViewSchedule schedule, bool numbersAsStrings)
        {
            TableData tableData = schedule.GetTableData();
            TableSectionData sectionData = tableData.GetSectionData(SectionType.Body);

            int rowCount = sectionData.NumberOfRows;
            int colCount = sectionData.NumberOfColumns;

            var data = new List<List<object>>(rowCount);

            for (int row = 0; row < rowCount; row++)
            {
                var rowData = new List<object>(colCount);

                for (int col = 0; col < colCount; col++)
                {
                    string cellValue = sectionData.GetCellText(row, col).Trim();

                    if (!numbersAsStrings)
                    {
                        // Attempt to convert cell value to numeric types
                        if (double.TryParse(cellValue, out double doubleValue))
                        {
                            rowData.Add(doubleValue);
                        }
                        else if (int.TryParse(cellValue, out int intValue))
                        {
                            rowData.Add(intValue);
                        }
                        else
                        {
                            rowData.Add(cellValue);
                        }
                    }
                    else
                    {
                        rowData.Add(cellValue);
                    }
                }

                data.Add(rowData);
            }

            return data;
        }


        private static string SanitizeWorksheetName(string name)
        {
            // Replace characters that are not supported in sheet names
            string sanitized = Regex.Replace(name, @"[:\\/[\]*?]", "");

            return sanitized;
        }

        private static bool IsFileOpen(string filePath)
        {
            try
            {
                using FileStream stream = new(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                stream.Close();
            }
            catch (IOException)
            {
                return true;
            }
            return false;
        }
    }
}
