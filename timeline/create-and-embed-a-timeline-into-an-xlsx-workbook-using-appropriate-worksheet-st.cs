using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data with dates and sales values
        cells["A1"].Value = "Date";
        cells["B1"].Value = "Sales";

        cells["A2"].Value = new DateTime(2023, 1, 1);
        cells["B2"].Value = 1000;
        cells["A3"].Value = new DateTime(2023, 1, 15);
        cells["B3"].Value = 1500;
        cells["A4"].Value = new DateTime(2023, 2, 1);
        cells["B4"].Value = 2000;
        cells["A5"].Value = new DateTime(2023, 2, 15);
        cells["B5"].Value = 2500;

        // Create a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        // Add Date as row field and Sales as data field
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        // Refresh and calculate the pivot table
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a timeline linked to the pivot table.
        // Placed at row 10, column 0 (cell A11)
        int timelineIndex = sheet.Timelines.Add(pivot, 10, 0, "Date");
        Timeline timeline = sheet.Timelines[timelineIndex];

        // Configure timeline appearance (optional)
        timeline.Caption = "Sales Timeline";
        timeline.Name = "SalesTimeline";
        timeline.LeftPixel = 50;   // horizontal offset from left column
        timeline.TopPixel = 20;    // vertical offset from top row
        timeline.WidthPixel = 400; // width in pixels
        timeline.HeightPixel = 80; // height in pixels

        // Save the workbook to an XLSX file
        workbook.Save("TimelineDemo.xlsx");
    }
}