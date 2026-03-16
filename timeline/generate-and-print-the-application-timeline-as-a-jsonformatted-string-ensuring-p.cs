using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate worksheet with sample data that includes a date field
        cells["A1"].PutValue("Date");
        cells["B1"].PutValue("Value");
        cells["A2"].PutValue(DateTime.Now.AddDays(-2));
        cells["B2"].PutValue(100);
        cells["A3"].PutValue(DateTime.Now.AddDays(-1));
        cells["B3"].PutValue(200);
        cells["A4"].PutValue(DateTime.Now);
        cells["B4"].PutValue(300);

        // Add a pivot table that will serve as the data source for the timeline
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "Pivot1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a timeline linked to the pivot table
        sheet.Timelines.Add(pivot, "F1", "Date");

        // Configure JSON save options with indentation for pretty printing
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            Indent = "    ",               // 4 spaces
            HasHeaderRow = true,          // treat first row as header
            ExportNestedStructure = true  // export as nested JSON structure
        };

        // Determine the used range of the worksheet
        int maxRow = sheet.Cells.MaxDataRow;
        int maxCol = sheet.Cells.MaxDataColumn;
        AsposeRange usedRange = sheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);

        // Convert the range to a JSON-formatted string using the options
        string json = usedRange.ToJson(jsonOptions);

        // Output the JSON string
        Console.WriteLine(json);
    }
}