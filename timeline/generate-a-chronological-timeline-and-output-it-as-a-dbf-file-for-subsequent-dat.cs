using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Saving;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate worksheet with chronological data (Date + Event)
        cells["A1"].PutValue("Date");
        cells["B1"].PutValue("Event");

        cells["A2"].PutValue(new DateTime(2023, 1, 1));
        cells["B2"].PutValue("New Year");

        cells["A3"].PutValue(new DateTime(2023, 2, 14));
        cells["B3"].PutValue("Valentine's Day");

        cells["A4"].PutValue(new DateTime(2023, 3, 17));
        cells["B4"].PutValue("St. Patrick's Day");

        cells["A5"].PutValue(new DateTime(2023, 4, 1));
        cells["B5"].PutValue("April Fool's Day");

        // Create a pivot table that uses the Date field as the row field
        PivotTableCollection pivots = sheet.PivotTables;
        int pivotIndex = pivots.Add("A1:B5", "D1", "TimelinePivot");
        PivotTable pivot = pivots[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Event");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a Timeline control linked to the pivot table, positioned at cell E1
        sheet.Timelines.Add(pivot, "E1", "Date");

        // Configure DBF save options (export all values as strings)
        DbfSaveOptions saveOptions = new DbfSaveOptions();
        saveOptions.ExportAsString = true;

        // Save the workbook as a DBF file
        workbook.Save("ChronologicalTimeline.dbf", saveOptions);
    }
}