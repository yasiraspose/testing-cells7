using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineToHtmlDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate worksheet with sample date/value data
        cells["A1"].Value = "Date";
        cells["B1"].Value = "Value";

        DateTime baseDate = new DateTime(2023, 1, 1);
        for (int i = 0; i < 10; i++)
        {
            // Add a date (monthly) and a corresponding numeric value
            cells[i + 1, 0].Value = baseDate.AddMonths(i);
            cells[i + 1, 1].Value = 50 + i * 10;
        }

        // Create a pivot table that will serve as the data source for the timeline
        int pivotIndex = sheet.PivotTables.Add("A1:B11", "D1", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a timeline linked to the pivot table's date field
        int timelineIndex = sheet.Timelines.Add(pivot, 0, 0, "Date");
        Timeline timeline = sheet.Timelines[timelineIndex];
        timeline.Caption = "Sales Timeline";

        // Save the workbook as HTML; the timeline will be rendered as part of the sheet
        workbook.Save("Timeline.html", SaveFormat.Html);
    }
}