using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineFromCsv
{
    class Program
    {
        static void Main()
        {
            // Path to the CSV file (replace with actual path)
            string csvPath = "data.csv";

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Import CSV data starting at cell A1 (row 0, column 0)
            // Using comma as delimiter and converting numeric data
            cells.ImportCSV(csvPath, ",", true, 0, 0);

            // Determine the used range after import
            int lastRow = cells.MaxDataRow;
            int lastCol = cells.MaxDataColumn;
            string startCell = CellsHelper.CellIndexToName(0, 0);
            string endCell = CellsHelper.CellIndexToName(lastRow, lastCol);
            string sourceRange = $"{startCell}:{endCell}";

            // Add a pivot table using the imported data
            // Destination cell for the pivot table is set to C1 (row 0, column 2)
            int pivotIndex = sheet.PivotTables.Add(sourceRange, "C1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            // Assuming the CSV has a column named "Date" and a column named "Value"
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh and calculate the pivot table data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table
            // Place the timeline at the top-left corner of the worksheet (row 0, column 0)
            int timelineIndex = sheet.Timelines.Add(pivot, 0, 0, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Optional: set a name for the timeline
            timeline.Name = "SalesTimeline";

            // Save the workbook to an XLSX file
            workbook.Save("TimelineFromCsv.xlsx");
        }
    }
}