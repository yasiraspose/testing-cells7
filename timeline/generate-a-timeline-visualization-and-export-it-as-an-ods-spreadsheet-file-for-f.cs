using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineToOds
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // ----- Populate sample data (Date, Category, Amount) -----
        cells["A1"].Value = "Date";
        cells["B1"].Value = "Category";
        cells["C1"].Value = "Amount";

        cells["A2"].Value = new DateTime(2023, 1, 1);
        cells["B2"].Value = "Alpha";
        cells["C2"].Value = 120;

        cells["A3"].Value = new DateTime(2023, 1, 5);
        cells["B3"].Value = "Beta";
        cells["C3"].Value = 150;

        cells["A4"].Value = new DateTime(2023, 1, 10);
        cells["B4"].Value = "Alpha";
        cells["C4"].Value = 200;

        cells["A5"].Value = new DateTime(2023, 1, 15);
        cells["B5"].Value = "Beta";
        cells["C5"].Value = 180;

        // ----- Create a PivotTable based on the data range -----
        int pivotIdx = sheet.PivotTables.Add("A1:C5", "E1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");          // Date as row field
        pivot.AddFieldToArea(PivotFieldType.Column, "Category");  // Category as column field
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");      // Amount as data field
        pivot.RefreshData();
        pivot.CalculateData();

        // ----- Add a Timeline linked to the PivotTable -----
        // Use the Add(PivotTable, string, string) overload (destCellName, baseFieldName)
        int timelineIdx = sheet.Timelines.Add(pivot, "A20", "Date");
        Timeline timeline = sheet.Timelines[timelineIdx];
        timeline.Caption = "Sales Timeline";

        // ----- Save the workbook as an ODS file -----
        workbook.Save("TimelineVisualization.ods", SaveFormat.Ods);
    }
}