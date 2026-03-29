using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;

class TimelineHtmlExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data with a date column and a numeric column
        cells["A1"].Value = "Date";
        cells["B1"].Value = "Value";
        cells["A2"].Value = new DateTime(2023, 1, 1);
        cells["B2"].Value = 100;
        cells["A3"].Value = new DateTime(2023, 2, 1);
        cells["B3"].Value = 150;
        cells["A4"].Value = new DateTime(2023, 3, 1);
        cells["B4"].Value = 200;

        // Create a pivot table that will serve as the data source for the timeline
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "Pivot1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a timeline linked to the pivot table (placed at cell A1)
        int timelineIdx = sheet.Timelines.Add(pivot, 0, 0, "Date");
        Timeline timeline = sheet.Timelines[timelineIdx];

        // Insert HTML‑formatted text into the timeline shape.
        // The HtmlText property accepts any valid HTML markup.
        timeline.Shape.HtmlText = "<div style='font-family:Arial;color:#2E86C1;'><b>Sales Timeline</b></div>";

        // Save the workbook as HTML, preserving the HTML markup inside the timeline.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportImagesAsBase64 = true;          // embed images (including timeline) as Base64
        htmlOptions.ParseHtmlTagInCell = true;           // ensure HTML tags are parsed
        htmlOptions.ExportHiddenWorksheet = true;        // include hidden sheets if any
        htmlOptions.ExportActiveWorksheetOnly = false;   // export the whole workbook
        workbook.Save("TimelineHtmlOutput.html", htmlOptions);
    }
}