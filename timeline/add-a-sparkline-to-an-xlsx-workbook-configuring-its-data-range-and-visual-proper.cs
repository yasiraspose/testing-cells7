using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline (row 0, columns A‑D)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell area where the sparkline will be placed (cell E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4, // Column E (0‑based index)
            EndColumn = 4
        };

        // Add a sparkline group of type Line, using the data range A1:D1,
        // plotted horizontally (isVertical = false), and positioned at the defined location
        int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIdx];

        // Add a sparkline to the group at row 0, column 4 (cell E1)
        int sparkIdx = group.Sparklines.Add("A1:D1", 0, 4);
        Sparkline spark = group.Sparklines[sparkIdx];

        // ----- Visual customizations -----

        // Set the series (line) color
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = Color.Orange;
        group.SeriesColor = seriesColor;

        // Highlight high and low points with custom colors
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        CellsColor highColor = workbook.CreateCellsColor();
        highColor.Color = Color.Green;
        group.HighPointColor = highColor;

        CellsColor lowColor = workbook.CreateCellsColor();
        lowColor.Color = Color.Red;
        group.LowPointColor = lowColor;

        // Adjust line thickness
        group.LineWeight = 1.0;

        // Apply a preset style (optional)
        group.PresetStyle = SparklinePresetStyleType.Style5;

        // Save the workbook to an XLSX file
        workbook.Save("SparklineDemo.xlsx");
    }
}