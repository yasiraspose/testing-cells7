using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data (Date and Value columns)
        cells["A1"].PutValue("Date");
        cells["B1"].PutValue("Value");

        DateTime[] dates = {
            new DateTime(2021, 1, 1),
            new DateTime(2021, 2, 1),
            new DateTime(2021, 3, 1)
        };
        int[] values = { 10, 20, 30 };

        for (int i = 0; i < dates.Length; i++)
        {
            cells[i + 1, 0].PutValue(dates[i]);
            cells[i + 1, 1].PutValue(values[i]);
        }

        // Create a PivotTable that will serve as the data source for the Timeline
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "Pivot1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a Timeline linked to the PivotTable
        int timelineIndex = sheet.Timelines.Add(pivot, 0, 0, "Date");
        Timeline timeline = sheet.Timelines[timelineIndex];
        timeline.Caption = "Demo Timeline";

        // Render the source data as tab‑delimited text, preserving order
        List<string> lines = new List<string>();
        int totalRows = cells.MaxDataRow + 1; // include header row
        for (int row = 0; row < totalRows; row++)
        {
            object dateObj = cells[row, 0].Value;
            object valueObj = cells[row, 1].Value;
            string line = $"{dateObj}\t{valueObj}";
            lines.Add(line);
        }

        string tabDelimitedOutput = string.Join(Environment.NewLine, lines);
        Console.WriteLine(tabDelimitedOutput);

        // Optionally save the same output to a .txt file using TxtSaveOptions
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            Separator = '\t',
            FormatStrategy = CellValueFormatStrategy.DisplayStyle
        };
        workbook.Save("TimelineData.txt", txtOptions);
    }
}