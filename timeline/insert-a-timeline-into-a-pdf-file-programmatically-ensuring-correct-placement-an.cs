using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

class InsertTimelineToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate worksheet with sample date and sales data
        sheet.Cells["A1"].Value = "Date";
        sheet.Cells["B1"].Value = "Sales";
        sheet.Cells["A2"].Value = new DateTime(2023, 1, 1);
        sheet.Cells["B2"].Value = 100;
        sheet.Cells["A3"].Value = new DateTime(2023, 2, 1);
        sheet.Cells["B3"].Value = 150;
        sheet.Cells["A4"].Value = new DateTime(2023, 3, 1);
        sheet.Cells["B4"].Value = 200;

        // Add a pivot table that will serve as the data source for the timeline
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.RefreshData();
        pivot.CalculateData();

        // Insert a timeline linked to the pivot table.
        // Row and column indices are zero‑based; (9,2) corresponds to cell C10.
        int timelineIndex = sheet.Timelines.Add(pivot, 9, 2, "Date");
        Timeline timeline = sheet.Timelines[timelineIndex];

        // Adjust visual appearance of the timeline via its Shape object
        timeline.Shape.Width = 400;   // width in pixels
        timeline.Shape.Height = 80;   // height in pixels
        timeline.Shape.Top = 200;     // vertical offset in pixels
        timeline.Shape.Left = 100;    // horizontal offset in pixels
        timeline.Caption = "Sales Timeline";

        // Save the workbook as a PDF, ensuring the timeline appears on a single page
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true
        };
        workbook.Save("TimelineOutput.pdf", pdfOptions);
    }
}