using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineDemo
{
    static void Main()
    {
        // Load the existing XLS workbook
        Workbook workbook = new Workbook("input.xls");
        Worksheet sheet = workbook.Worksheets[0];

        // ------------------------------------------------------------
        // Create a PivotTable that will serve as the data source for the Timeline.
        // Assume the worksheet already contains date values in column A and
        // corresponding numeric values (e.g., sales) in column B.
        // ------------------------------------------------------------
        // Add the PivotTable covering a reasonable range (adjust as needed)
        int pivotIndex = sheet.PivotTables.Add("A1:B100", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Configure the PivotTable: Date as Row field, Sales as Data field
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Refresh the PivotTable so it reflects the source data
        pivot.RefreshData();
        pivot.CalculateData();

        // ------------------------------------------------------------
        // Add a Timeline control linked to the PivotTable.
        // The Timeline will be placed with its upper‑left corner at cell E1
        // and will use the "Date" field as its base field.
        // ------------------------------------------------------------
        sheet.Timelines.Add(pivot, "E1", "Date");

        // Save the workbook with the newly added Timeline
        workbook.Save("output.xls");
    }
}