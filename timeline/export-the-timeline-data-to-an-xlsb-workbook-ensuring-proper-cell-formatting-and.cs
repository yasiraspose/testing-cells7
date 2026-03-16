using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineExport
{
    public class ExportTimelineToXlsb
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate worksheet with sample data
            // Header row
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Date";
            cells["C1"].Value = "Amount";

            // Sample rows
            cells["A2"].Value = "Fruit";
            cells["B2"].Value = new DateTime(2021, 2, 5);
            cells["C2"].Value = 50;

            cells["A3"].Value = "Fruit";
            cells["B3"].Value = new DateTime(2022, 3, 8);
            cells["C3"].Value = 60;

            cells["A4"].Value = "Fruit";
            cells["B4"].Value = new DateTime(2023, 4, 10);
            cells["C4"].Value = 70;

            cells["A5"].Value = "Fruit";
            cells["B5"].Value = new DateTime(2024, 5, 16);
            cells["C5"].Value = 80;

            // Create a date style and apply it to the date column
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "m/d/yyyy";
            for (int row = 1; row <= 4; row++)
            {
                cells[row, 1].SetStyle(dateStyle);
            }

            // Add a PivotTable based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Column, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium10;

            // Refresh and calculate the PivotTable data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline control linked to the PivotTable
            // Place the Timeline starting at row 10, column 5 (cell E10)
            int timelineIndex = sheet.Timelines.Add(pivot, 9, 4, "Date"); // zero‑based indices
            Timeline timeline = sheet.Timelines[timelineIndex];
            timeline.Caption = "Sales Timeline";
            timeline.ShowHeader = true;
            timeline.ShowHorizontalScrollbar = true;
            timeline.ShowSelectionLabel = true;
            timeline.ShowTimeLevel = true;

            // Configure XLSB save options
            XlsbSaveOptions saveOptions = new XlsbSaveOptions
            {
                ExportAllColumnIndexes = true // ensure all column indexes are exported
            };

            // Save the workbook as an XLSB file
            workbook.Save("TimelineData.xlsb", saveOptions);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportTimelineToXlsb.Run();
        }
    }
}