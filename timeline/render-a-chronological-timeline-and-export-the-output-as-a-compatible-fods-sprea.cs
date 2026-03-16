using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace TimelineExportDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate sample data (fruit, date, amount)
            cells[0, 0].Value = "fruit";
            cells[0, 1].Value = "date";
            cells[0, 2].Value = "amount";

            cells[1, 0].Value = "grape";
            cells[2, 0].Value = "blueberry";
            cells[3, 0].Value = "kiwi";
            cells[4, 0].Value = "cherry";

            // Dates for the timeline
            cells[1, 1].Value = new DateTime(2021, 2, 5);
            cells[2, 1].Value = new DateTime(2022, 3, 8);
            cells[3, 1].Value = new DateTime(2023, 4, 10);
            cells[4, 1].Value = new DateTime(2024, 5, 16);

            // Amounts
            cells[1, 2].Value = 50;
            cells[2, 2].Value = 60;
            cells[3, 2].Value = 70;
            cells[4, 2].Value = 80;

            // 3. Create a PivotTable based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("A1:C5", "E1", "FruitPivot");
            PivotTable pivot = pivots[pivotIndex];

            // Add fields to the PivotTable
            pivot.AddFieldToArea(PivotFieldType.Row, "fruit");
            pivot.AddFieldToArea(PivotFieldType.Column, "date");
            pivot.AddFieldToArea(PivotFieldType.Data, "amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium10;

            // Refresh and calculate the PivotTable
            pivot.RefreshData();
            pivot.CalculateData();

            // 4. Add a Timeline control linked to the PivotTable's date field
            // Place the Timeline starting at cell A20
            TimelineCollection timelines = sheet.Timelines;
            int timelineIndex = timelines.Add(pivot, "A20", "date");
            Timeline timeline = timelines[timelineIndex];

            // Optional: set a caption for the timeline
            timeline.Caption = "Fruit Sales Timeline";

            // 5. Save the workbook as an OpenDocument Flat XML Spreadsheet (FODS)
            workbook.Save("FruitTimeline.fods", SaveFormat.Fods);
        }
    }
}