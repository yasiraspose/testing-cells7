using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineMhtDemo
{
    static void Main()
    {
        // 1. Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // 2. Populate worksheet with sample data (date and sales)
        cells["A1"].Value = "Date";
        cells["B1"].Value = "Sales";

        cells["A2"].Value = new DateTime(2023, 1, 1);
        cells["B2"].Value = 1200;

        cells["A3"].Value = new DateTime(2023, 2, 1);
        cells["B3"].Value = 1500;

        cells["A4"].Value = new DateTime(2023, 3, 1);
        cells["B4"].Value = 1800;

        // 3. Add a PivotTable that will serve as the data source for the Timeline
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.RefreshData();
        pivot.CalculateData();

        // 4. Add a Timeline linked to the PivotTable (using the "Date" field)
        int timelineIdx = sheet.Timelines.Add(pivot, 0, 0, "Date");
        Timeline timeline = sheet.Timelines[timelineIdx];
        timeline.Caption = "Sales Timeline";
        timeline.Name = "SalesTimeline";
        timeline.Shape.Width = 400;   // width in pixels
        timeline.Shape.Height = 80;   // height in pixels

        // 5. Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.PresentationPreference = true;
        saveOptions.ExportFrameScriptsAndProperties = true;

        // 6. Save the workbook as an HTML file
        workbook.Save("TimelineDemo.html", saveOptions);
    }
}