using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineToMhtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate worksheet with sample data
        cells[0, 0].PutValue("Fruit");
        cells[1, 0].PutValue("Grape");
        cells[2, 0].PutValue("Blueberry");
        cells[3, 0].PutValue("Kiwi");
        cells[4, 0].PutValue("Cherry");

        // Create a date style
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "m/d/yyyy";

        cells[0, 1].PutValue("Date");
        cells[1, 1].PutValue(new DateTime(2021, 2, 5));
        cells[2, 1].PutValue(new DateTime(2022, 3, 8));
        cells[3, 1].PutValue(new DateTime(2023, 4, 10));
        cells[4, 1].PutValue(new DateTime(2024, 5, 16));
        cells[1, 1].SetStyle(dateStyle);
        cells[2, 1].SetStyle(dateStyle);
        cells[3, 1].SetStyle(dateStyle);
        cells[4, 1].SetStyle(dateStyle);

        cells[0, 2].PutValue("Amount");
        cells[1, 2].PutValue(50);
        cells[2, 2].PutValue(60);
        cells[3, 2].PutValue(70);
        cells[4, 2].PutValue(80);

        // Add a PivotTable based on the data range
        PivotTableCollection pivots = sheet.PivotTables;
        int pivotIndex = pivots.Add("=Sheet1!A1:C5", "A12", "SalesPivot");
        PivotTable pivot = pivots[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "fruit");
        pivot.AddFieldToArea(PivotFieldType.Column, "date");
        pivot.AddFieldToArea(PivotFieldType.Data, "amount");
        pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium10;
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a Timeline linked to the PivotTable's date field
        sheet.Timelines.Add(pivot, 10, 5, "date");
        Timeline timeline = sheet.Timelines[0];
        timeline.Caption = "Sales Timeline";
        timeline.ShowHeader = true;
        timeline.ShowHorizontalScrollbar = true;
        timeline.ShowSelectionLabel = true;
        timeline.ShowTimeLevel = true;

        // Configure HTML save options for MHTML output
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);
        saveOptions.PresentationPreference = true;               // Better visual presentation
        saveOptions.ExportFrameScriptsAndProperties = true;      // Preserve interactive elements
        saveOptions.ExportGridLines = true;                      // Keep grid lines
        saveOptions.IsMobileCompatible = false;                 // Not targeting mobile specifically

        // Save the workbook as an MHTML file
        workbook.Save("TimelineOutput.mht", saveOptions);
    }
}