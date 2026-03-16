// Import necessary Aspose.Cells namespaces
using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsTimelineFromFods
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook instance
            Workbook workbook = new Workbook();

            // 2. Import the FODS (OpenDocument Spreadsheet) XML data into the first worksheet.
            //    The ImportXml method treats the FODS file as an XML source.
            //    Adjust the file path, target sheet name, and start cell as needed.
            string fodsFilePath = "input.fods";          // Path to the source FODS file
            string targetSheetName = "Sheet1";           // Destination sheet name
            int startRow = 0;                            // Zero‑based row index (A1)
            int startColumn = 0;                         // Zero‑based column index (A1)

            workbook.ImportXml(fodsFilePath, targetSheetName, startRow, startColumn);

            // 3. Access the worksheet that now contains the imported data
            Worksheet sheet = workbook.Worksheets[0];

            // 4. Define the data range for the pivot table.
            //    Here we assume the imported data occupies columns A‑C and rows 1‑5.
            //    Adjust the range string if the actual data size differs.
            string dataRange = "A1:C5";

            // 5. Add a pivot table to the worksheet.
            //    The pivot table will be placed starting at cell E3 with the name "PivotTable1".
            int pivotIndex = sheet.PivotTables.Add(dataRange, "E3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // 6. Configure the pivot table fields.
            //    Assume the first column is a date field named "Date".
            //    Adjust field names according to the actual imported columns.
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");          // Date as row field
            pivot.AddFieldToArea(PivotFieldType.Data, 1);             // Second column as data field

            // 7. Refresh the pivot table to ensure it reflects the imported data.
            pivot.RefreshData();
            pivot.CalculateData();

            // 8. Add a Timeline control linked to the pivot table.
            //    Use the overload that specifies the destination cell name and the base field name.
            //    The timeline will be placed with its upper‑left corner at cell G1.
            string timelineDestCell = "G1";
            string baseFieldName = "Date";   // The field in the pivot's BaseFields to drive the timeline
            sheet.Timelines.Add(pivot, timelineDestCell, baseFieldName);

            // 9. Save the resulting workbook to an XLSX file.
            string outputPath = "TimelineFromFods.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook with timeline saved to: {outputPath}");
        }
    }
}