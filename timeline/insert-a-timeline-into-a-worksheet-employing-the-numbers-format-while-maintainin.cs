using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineExample
{
    public class InsertTimelineWithNumberFormat
    {
        public static void Run()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate worksheet with sample data (including a date column)
            // Header row
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Date");
            cells["C1"].PutValue("Sales");

            // Sample rows
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(new DateTime(2023, 1, 1));
            cells["C2"].PutValue(1200);

            cells["A3"].PutValue("B");
            cells["B3"].PutValue(new DateTime(2023, 2, 1));
            cells["C3"].PutValue(1500);

            cells["A4"].PutValue("C");
            cells["B4"].PutValue(new DateTime(2023, 3, 1));
            cells["C4"].PutValue(1800);

            // 3. Apply a numeric (date) format to the Date column to ensure proper display
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "m/d/yyyy"; // Number format for dates
            cells["B2"].SetStyle(dateStyle);
            cells["B3"].SetStyle(dateStyle);
            cells["B4"].SetStyle(dateStyle);

            // 4. Insert a pivot table that will serve as the data source for the timeline
            // Place the pivot table starting at cell E2 to keep it away from existing data
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E2", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            // Add fields: Date as row, Sales as data
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // 5. Add a timeline linked to the pivot table.
            // Position the timeline at cell H2 (far enough not to overwrite existing content)
            int timelineIndex = sheet.Timelines.Add(pivot, "H2", "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Optional: set a caption for clarity
            timeline.Caption = "Sales Timeline";

            // 6. Save the workbook (output file)
            workbook.Save("TimelineWithNumberFormat.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            InsertTimelineWithNumberFormat.Run();
        }
    }
}