using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;

namespace TimelineTemplateDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate worksheet with sample data (fruit, date, amount)
            cells[0, 0].Value = "fruit";
            cells[1, 0].Value = "grape";
            cells[2, 0].Value = "blueberry";
            cells[3, 0].Value = "kiwi";
            cells[4, 0].Value = "cherry";

            // Create a date style
            Style dateStyle = new CellsFactory().CreateStyle();
            dateStyle.Custom = "m/d/yyyy";

            cells[0, 1].Value = "date";
            cells[1, 1].Value = new DateTime(2021, 2, 5);
            cells[2, 1].Value = new DateTime(2022, 3, 8);
            cells[3, 1].Value = new DateTime(2023, 4, 10);
            cells[4, 1].Value = new DateTime(2024, 5, 16);
            cells[1, 1].SetStyle(dateStyle);
            cells[2, 1].SetStyle(dateStyle);
            cells[3, 1].SetStyle(dateStyle);
            cells[4, 1].SetStyle(dateStyle);

            cells[0, 2].Value = "amount";
            cells[1, 2].Value = 50;
            cells[2, 2].Value = 60;
            cells[3, 2].Value = 70;
            cells[4, 2].Value = 80;

            // Add a PivotTable based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("=Sheet1!A1:C5", "A12", "TestPivotTable");
            PivotTable pivot = pivots[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "fruit");
            pivot.AddFieldToArea(PivotFieldType.Column, "date");
            pivot.AddFieldToArea(PivotFieldType.Data, "amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium10;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the PivotTable (using the rule Add(PivotTable, string, string))
            TimelineCollection timelines = sheet.Timelines;
            int timelineIndex = timelines.Add(pivot, "A20", "date");
            Timeline timeline = timelines[timelineIndex];

            // Set some Timeline properties (caption, name, position, size)
            timeline.Caption = "Fruit Sales Timeline";
            timeline.Name = "FruitTimeline";
            timeline.LeftPixel = 100;   // obsolete property, kept for compatibility
            timeline.TopPixel = 50;     // obsolete property, kept for compatibility
            timeline.WidthPixel = 400;  // obsolete property, kept for compatibility
            timeline.HeightPixel = 120; // obsolete property, kept for compatibility

            // Save the workbook as an XLTM template (using XlsSaveOptions.IsTemplate rule)
            XlsSaveOptions saveOptions = new XlsSaveOptions();
            saveOptions.IsTemplate = true; // marks the file as a template
            workbook.Save("FruitTimelineTemplate.xltm", saveOptions);
        }
    }
}