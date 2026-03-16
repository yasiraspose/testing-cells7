using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate worksheet with sample date and value data
        cells["A1"].PutValue("Date");
        cells["B1"].PutValue("Value");
        cells["A2"].PutValue(new DateTime(2021, 1, 1));
        cells["B2"].PutValue(100);
        cells["A3"].PutValue(new DateTime(2021, 2, 1));
        cells["B3"].PutValue(150);
        cells["A4"].PutValue(new DateTime(2021, 3, 1));
        cells["B4"].PutValue(200);

        // Add a PivotTable that will serve as the data source for the Timeline
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a Timeline linked to the PivotTable using the "Date" field
        int timelineIndex = sheet.Timelines.Add(pivot, "F1", "Date");
        Timeline timeline = sheet.Timelines[timelineIndex];
        timeline.Caption = "Sales Timeline";

        // Save the workbook in StarOffice Calc (SXC) format
        workbook.Save("TimelineOutput.sxc", SaveFormat.Sxc);
    }
}