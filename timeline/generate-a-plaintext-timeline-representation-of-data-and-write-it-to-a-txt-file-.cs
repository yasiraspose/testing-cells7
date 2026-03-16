using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate worksheet with sample timeline data (Date and Value)
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");

        DateTime startDate = new DateTime(2021, 1, 1);
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[i + 1, 0].PutValue(startDate.AddMonths(i)); // Date column
            sheet.Cells[i + 1, 1].PutValue((i + 1) * 10);          // Value column
        }

        // Create a PivotTable that will serve as the data source for the Timeline
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();

        // Add a Timeline linked to the PivotTable (optional for visual representation)
        sheet.Timelines.Add(pivot, 0, 3, "Date");

        // Configure TxtSaveOptions to export the worksheet as plain‑text (tab‑delimited)
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            Separator = '\t',          // Use tab as column separator
            ExportAllSheets = false   // Export only the active sheet
        };

        // Save the workbook as a TXT file containing the timeline data
        workbook.Save("TimelineData.txt", txtOptions);
    }
}