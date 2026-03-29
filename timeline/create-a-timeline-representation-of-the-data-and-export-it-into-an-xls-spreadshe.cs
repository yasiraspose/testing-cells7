using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace TimelineExportExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: creation)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate worksheet with sample data (Date and Value columns)
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Value");

            // Sample dates and corresponding values
            DateTime[] dates = {
                new DateTime(2021, 1, 1),
                new DateTime(2021, 2, 1),
                new DateTime(2021, 3, 1),
                new DateTime(2021, 4, 1),
                new DateTime(2021, 5, 1)
            };
            int[] values = { 100, 150, 130, 170, 160 };

            for (int i = 0; i < dates.Length; i++)
            {
                cells[i + 1, 0].PutValue(dates[i]);   // Column A: Date
                cells[i + 1, 1].PutValue(values[i]); // Column B: Value
            }

            // Create a pivot table that will serve as the data source for the timeline
            // The pivot range includes both columns; the destination cell is C3
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "C3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add the Date field to the Row area and the Value field to the Data area
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh and calculate the pivot table data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table.
            // Place the timeline starting at row 10, column 0 (cell A10) and bind it to the "Date" field.
            int timelineIndex = sheet.Timelines.Add(pivot, 9, 0, "Date"); // Row and column are zero‑based
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Optional: set some visual properties of the timeline
            timeline.Caption = "Sales Timeline";
            timeline.ShowHeader = true;
            timeline.ShowHorizontalScrollbar = true;

            // Save the workbook as an XLS file (lifecycle: saving)
            workbook.Save("TimelineOutput.xls");
        }
    }
}