using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineDifExport
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate data with hierarchy (Category, SubCategory) and timestamps
        cells["A1"].Value = "Category";
        cells["B1"].Value = "SubCategory";
        cells["C1"].Value = "Date";
        cells["D1"].Value = "Value";

        cells["A2"].Value = "Fruit";
        cells["B2"].Value = "Apple";
        cells["C2"].Value = new DateTime(2023, 1, 1);
        cells["D2"].Value = 120;

        cells["A3"].Value = "Fruit";
        cells["B3"].Value = "Apple";
        cells["C3"].Value = new DateTime(2023, 1, 15);
        cells["D3"].Value = 150;

        cells["A4"].Value = "Fruit";
        cells["B4"].Value = "Banana";
        cells["C4"].Value = new DateTime(2023, 2, 5);
        cells["D4"].Value = 200;

        cells["A5"].Value = "Vegetable";
        cells["B5"].Value = "Carrot";
        cells["C5"].Value = new DateTime(2023, 3, 10);
        cells["D5"].Value = 80;

        // Create a pivot table using the data range A1:D5, place it at F1
        int pivotIndex = sheet.PivotTables.Add("A1:D5", "F1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Configure pivot fields: Category & SubCategory as rows, Date as column, Value as data
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Row, "SubCategory");
        pivot.AddFieldToArea(PivotFieldType.Column, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");

        // Refresh and calculate the pivot table
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a timeline linked to the Date field of the pivot table
        // Position the timeline at row 12, column 0 (cell A13)
        int timelineIndex = sheet.Timelines.Add(pivot, 12, 0, "Date");
        Timeline timeline = sheet.Timelines[timelineIndex];
        timeline.Caption = "Sales Timeline";

        // Save the workbook in DIF format using DifSaveOptions
        DifSaveOptions difOptions = new DifSaveOptions()
        {
            ClearData = true,
            CreateDirectory = true,
            RefreshChartCache = true
        };

        workbook.Save("TimelineOutput.dif", difOptions);
    }
}