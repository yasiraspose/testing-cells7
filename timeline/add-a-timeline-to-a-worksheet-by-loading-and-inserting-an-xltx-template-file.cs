using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineExample
{
    class Program
    {
        static void Main()
        {
            // Load the XLTX template file (replace with actual template path)
            Workbook workbook = new Workbook("Template.xltx");

            // Get the first worksheet from the loaded workbook
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Ensure the worksheet contains the required data for a pivot table.
            // This example assumes the template already has a date column (A) 
            // and a numeric column (B). If not, you can uncomment the lines 
            // below to add sample data.
            // ------------------------------------------------------------
            //cells["A1"].PutValue("Date");
            //cells["B1"].PutValue("Sales");
            //cells["A2"].PutValue(new DateTime(2023, 1, 1));
            //cells["B2"].PutValue(1000);
            //cells["A3"].PutValue(new DateTime(2023, 2, 1));
            //cells["B3"].PutValue(1500);
            //cells["A4"].PutValue(new DateTime(2023, 3, 1));
            //cells["B4"].PutValue(2000);

            // Create a pivot table that will serve as the data source for the timeline
            // The data range covers columns A and B, rows 1 through 4 (adjust as needed)
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table: Date as row field, Sales as data field
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table to ensure it has data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table.
            // The timeline will be placed with its upper‑left corner at cell F1
            // and will use the "Date" field from the pivot table.
            int timelineIndex = sheet.Timelines.Add(pivot, "F1", "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Optional: set a friendly name for the timeline
            timeline.Name = "SalesTimeline";

            // Save the workbook with the newly added timeline
            workbook.Save("Result.xlsx");
        }
    }
}