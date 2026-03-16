using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelinePrnDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate worksheet with sample data (date and value)
            cells["A1"].Value = "Date";
            cells["B1"].Value = "Value";

            cells["A2"].Value = new DateTime(2023, 1, 1);
            cells["B2"].Value = 100;

            cells["A3"].Value = new DateTime(2023, 2, 1);
            cells["B3"].Value = 150;

            cells["A4"].Value = new DateTime(2023, 3, 1);
            cells["B4"].Value = 200;

            // Create a pivot table that will serve as the data source for the timeline
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table (placed at cell A6)
            int timelineIndex = sheet.Timelines.Add(pivot, "A6", "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Optional: set some timeline properties
            timeline.Caption = "Sales Timeline";
            timeline.Name = "SalesTimeline";

            // Save the workbook
            workbook.Save("TimelineDemo.xlsx");
        }
    }
}