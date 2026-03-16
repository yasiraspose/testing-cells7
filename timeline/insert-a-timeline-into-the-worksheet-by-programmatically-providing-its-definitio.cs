using System;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsTimelineFromJson
{
    public class TimelineDefinition
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string BaseFieldName { get; set; } = null!;
    }

    public class Program
    {
        public static void Main()
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            cells["A1"].Value = "Date";
            cells["B1"].Value = "Sales";

            cells["A2"].Value = new DateTime(2023, 1, 1);
            cells["B2"].Value = 1200;

            cells["A3"].Value = new DateTime(2023, 2, 1);
            cells["B3"].Value = 1500;

            cells["A4"].Value = new DateTime(2023, 3, 1);
            cells["B4"].Value = 1800;

            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            string json = @"{
                ""Row"": 10,
                ""Column"": 5,
                ""BaseFieldName"": ""Date""
            }";

            TimelineDefinition definition = JsonSerializer.Deserialize<TimelineDefinition>(json)!;

            int timelineIndex = sheet.Timelines.Add(pivot, definition.Row, definition.Column, definition.BaseFieldName);
            var timeline = sheet.Timelines[timelineIndex];
            timeline.Caption = "Sales Timeline";

            workbook.Save("TimelineFromJson.xlsx");
        }
    }
}