using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineSpreadsheetMLDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate worksheet with sample chronological data
        cells[0, 0].Value = "Event";
        cells[0, 1].Value = "Date";
        cells[0, 2].Value = "Amount";

        cells[1, 0].Value = "Launch";
        cells[1, 1].Value = new DateTime(2021, 1, 15);
        cells[1, 2].Value = 1000;

        cells[2, 0].Value = "Update";
        cells[2, 1].Value = new DateTime(2022, 5, 10);
        cells[2, 2].Value = 1500;

        cells[3, 0].Value = "Release";
        cells[3, 1].Value = new DateTime(2023, 9, 30);
        cells[3, 2].Value = 2000;

        // Apply a date style to the date column
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "m/d/yyyy";
        cells[1, 1].SetStyle(dateStyle);
        cells[2, 1].SetStyle(dateStyle);
        cells[3, 1].SetStyle(dateStyle);

        // Create a PivotTable based on the data range and place it at E1
        PivotTableCollection pivots = sheet.PivotTables;
        int pivotIndex = pivots.Add("A1:C4", "E1", "TimelinePivot");
        PivotTable pivot = pivots[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Event");
        pivot.AddFieldToArea(PivotFieldType.Column, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a Timeline linked to the PivotTable, positioned at G10, using the "Date" field
        int timelineIndex = sheet.Timelines.Add(pivot, "G10", "Date");
        Timeline timeline = sheet.Timelines[timelineIndex];
        timeline.Caption = "Project Timeline";
        timeline.ShowHeader = true;

        // Save the workbook as an Excel 2003 XML (SpreadsheetML) file
        SpreadsheetML2003SaveOptions saveOptions = new SpreadsheetML2003SaveOptions();
        saveOptions.IsIndentedFormatting = true;          // optional: make XML indented for readability
        saveOptions.ExportColumnIndexOfCell = true;       // optional: include column index in cells
        workbook.Save("ProjectTimeline.xml", saveOptions);
    }
}