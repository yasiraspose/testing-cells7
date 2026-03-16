using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Loading;
using Aspose.Cells.Timelines;

class InsertTimelineFromDif
{
    static void Main()
    {
        // Load data from a DIF file using DifLoadOptions
        DifLoadOptions difOptions = new DifLoadOptions();
        Workbook workbook = new Workbook("input.dif", difOptions);

        // Get the first worksheet (assumed to contain the imported data)
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the used range of the data (assumes first row contains headers)
        int lastRow = sheet.Cells.MaxDataRow;          // zero‑based index of last row with data
        int lastCol = sheet.Cells.MaxDataColumn;       // zero‑based index of last column with data

        // Build the address string for the data range, e.g., "A1:B10"
        string startCell = CellsHelper.CellIndexToName(0, 0);                     // "A1"
        string endCell   = CellsHelper.CellIndexToName(lastRow, lastCol);        // e.g., "B10"
        string dataRange = $"{startCell}:{endCell}";

        // Add a pivot table that will serve as the data source for the timeline
        // Place the pivot table at cell "E1"
        int pivotIdx = sheet.PivotTables.Add(dataRange, "E1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];

        // Add the required fields to the pivot table.
        // Assumes the first column header is "Date" and the second is "Value".
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");

        // Refresh and calculate the pivot table data
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a timeline control linked to the pivot table.
        // The timeline will be placed with its upper‑left corner at cell "G1"
        // and will use the "Date" field as the base field.
        int timelineIdx = sheet.Timelines.Add(pivot, "G1", "Date");
        Timeline timeline = sheet.Timelines[timelineIdx];

        // Optional: set a caption for the timeline
        timeline.Caption = "Sales Timeline";

        // Save the workbook with the timeline added
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}