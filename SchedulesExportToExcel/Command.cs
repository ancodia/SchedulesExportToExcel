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
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
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
                ExportForm exportForm = new ExportForm(scheduleNames);
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

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                FileInfo excelFile = new FileInfo(filePath);
                if (excelFile.Exists)
                {
                    if (IsFileOpen(filePath))
                    {
                        MessageBox.Show("The Excel file is currently open. Please close it and try again.", "File In Use");
                        return;
                    }

                    // Load content of existing file if it exists
                    using (FileStream stream = new FileStream(filePath, FileMode.Open))
                    {
                        excelPackage.Load(stream);
                    }
                }

                foreach (ViewSchedule schedule in schedules)
                {
                    string scheduleName = SanitizeWorksheetName(schedule.Name);

                    // Delete existing worksheet if it exists
                    var worksheet = excelPackage.Workbook.Worksheets[scheduleName];
                    if (worksheet != null)
                    {
                        excelPackage.Workbook.Worksheets.Delete(scheduleName);
                    }

                    // Export the schedule to CSV
                    string tempDirectory = Path.GetTempPath();
                    string tempCsvFilePath = Path.Combine(tempDirectory, $"{scheduleName}.csv");

                    ViewScheduleExportOptions exportOptions = new ViewScheduleExportOptions();
                    exportOptions.Title = false;
                    exportOptions.FieldDelimiter = ",";
                    exportOptions.TextQualifier = ExportTextQualifier.None;
                    schedule.Export(tempDirectory, $"{scheduleName}.csv", exportOptions);

                    // Add new worksheet at the start of the excel workbook
                    worksheet = excelPackage.Workbook.Worksheets.Add(scheduleName);
                    excelPackage.Workbook.Worksheets.MoveToStart(scheduleName);

                    using (var reader = new StreamReader(tempCsvFilePath))
                    {
                        // Read all lines into an array
                        var lines = new List<string>();
                        while (!reader.EndOfStream)
                        {
                            lines.Add(reader.ReadLine());
                        }

                        // Prepare a list for batch writing
                        var rowCount = lines.Count;
                        var colCount = lines[0].Split(',').Length; // Assuming all rows have the same number of columns
                        var data = new List<object[]>(rowCount);

                        for (int row = 0; row < rowCount; row++)
                        {
                            var cells = lines[row].Split(',');
                            var rowData = new object[colCount];

                            for (int col = 0; col < colCount; col++)
                            {
                                string cellValue = cells[col].Trim();

                                if (!numbersAsStrings)
                                {
                                    // Attempt to convert cell values to numeric types
                                    if (double.TryParse(cellValue, out double doubleValue))
                                    {
                                        rowData[col] = doubleValue;
                                    }
                                    else if (int.TryParse(cellValue, out int intValue))
                                    {
                                        rowData[col] = intValue;
                                    }
                                    else
                                    {
                                        rowData[col] = cellValue;
                                    }
                                }
                                else
                                {
                                    rowData[col] = cellValue;
                                }
                            }

                            data.Add(rowData);
                        }

                        // Write the entire data list to the worksheet in one go
                        var startRow = 1;
                        var startCol = 1;

                        var range = worksheet.Cells[startRow, startCol, startRow + rowCount - 1, startCol + colCount - 1];
                        range.LoadFromArrays(data);
                    }

                    File.Delete(tempCsvFilePath);
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

        private string SanitizeWorksheetName(string name)
        {
            // Replace characters that are not supported in sheet names
            string sanitized = Regex.Replace(name, @"[:\\/[\]*?]", "");

            return sanitized;
        }

        private bool IsFileOpen(string filePath)
        {
            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                return true;
            }
            return false;
        }
    }
}
