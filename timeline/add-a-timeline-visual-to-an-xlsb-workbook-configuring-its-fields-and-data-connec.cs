using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSB will be set on save)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: Date column (A) and Sales column (B)
            cells["A1"].Value = "Date";
            cells["B1"].Value = "Sales";

            // Add a few rows of data
            cells["A2"].Value = new DateTime(2023, 1, 1);
            cells["B2"].Value = 1200;
            cells["A3"].Value = new DateTime(2023, 2, 1);
            cells["B3"].Value = 1500;
            cells["A4"].Value = new DateTime(2023, 3, 1);
            cells["B4"].Value = 1800;
            cells["A5"].Value = new DateTime(2023, 4, 1);
            cells["B5"].Value = 2100;

            // Create a pivot table that will serve as the data source for the timeline
            // Data range: A1:B5, destination cell for pivot table: D2, name: "SalesPivot"
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D2", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            // Row field: Date
            // Data field: Sales (sum)
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table.
            // Place the timeline's upper‑left corner at cell G2.
            // Base field name is the Date field used in the pivot.
            int timelineIndex = sheet.Timelines.Add(pivot, "G2", "Date");

            // Retrieve the created timeline to set additional properties (optional)
            Timeline timeline = sheet.Timelines[timelineIndex];
            timeline.Name = "SalesTimeline";
            timeline.Caption = "Sales Over Time";

            // Save the workbook as XLSB
            workbook.Save("TimelineDemo.xlsb", SaveFormat.Xlsb);
        }
    }
}