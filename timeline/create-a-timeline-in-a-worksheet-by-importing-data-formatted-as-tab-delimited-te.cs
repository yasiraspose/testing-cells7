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
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Import TAB‑delimited text file (ImportCSV with tab splitter)
            // Parameters: file name, splitter, convert numeric data, start row, start column
            string tsvPath = "data.tsv";               // Path to the TAB‑delimited file
            cells.ImportCSV(tsvPath, "\t", true, 0, 0);

            // Assume the imported data has two columns: "Date" (A) and "Sales" (B)
            // Create a pivot table using the imported range (A1:B5 as an example)
            // Adjust the range if your file has more rows
            string sourceRange = "A1:B5";
            string pivotDestCell = "D1";
            int pivotIndex = sheet.PivotTables.Add(sourceRange, pivotDestCell, "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");   // Date as row field
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales"); // Sales as data field

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline control linked to the pivot table, placed at cell E1
            sheet.Timelines.Add(pivot, "E1", "Date");

            // Save the workbook (lifecycle rule)
            workbook.Save("TimelineFromTsv.xlsx");
        }
    }
}