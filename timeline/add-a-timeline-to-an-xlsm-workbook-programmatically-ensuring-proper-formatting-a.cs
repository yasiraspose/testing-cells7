using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace TimelineXlsmDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSM will be set on save)
            Workbook workbook = new Workbook();

            // Enable macros for the workbook (required for XLSM compatibility)
            workbook.Settings.EnableMacros = true;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with a date column (required for timeline)
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Date";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "P1";
            cells["B2"].Value = new DateTime(2023, 1, 1);
            cells["C2"].Value = 120;

            cells["A3"].Value = "P2";
            cells["B3"].Value = new DateTime(2023, 1, 5);
            cells["C3"].Value = 150;

            cells["A4"].Value = "P1";
            cells["B4"].Value = new DateTime(2023, 1, 10);
            cells["C4"].Value = 200;

            // Create a pivot table that will serve as the timeline data source
            int pivotIdx = sheet.PivotTables.Add("A1:C4", "E1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Column, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

            // Refresh pivot data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table, positioned at cell A10
            int timelineIdx = sheet.Timelines.Add(pivot, "A10", "Date");
            Timeline timeline = sheet.Timelines[timelineIdx];

            // Configure timeline appearance
            timeline.Caption = "Sales Timeline";
            timeline.Name = "SalesTimeline";
            timeline.LeftPixel = 50;   // Horizontal offset
            timeline.TopPixel = 200;   // Vertical offset
            timeline.WidthPixel = 400; // Width
            timeline.HeightPixel = 120; // Height
            timeline.ShowHeader = true;
            timeline.ShowHorizontalScrollbar = true;
            timeline.ShowSelectionLabel = true;
            timeline.ShowTimeLevel = true;

            // Save as macro-enabled workbook
            workbook.Save("SalesTimelineDemo.xlsm", SaveFormat.Xlsm);
        }
    }
}