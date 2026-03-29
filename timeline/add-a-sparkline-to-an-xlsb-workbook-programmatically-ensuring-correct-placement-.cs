using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineXlsbExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (A1:C5)
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[i, 0].PutValue(i + 1);          // A column
            sheet.Cells[i, 1].PutValue((i + 1) * 2);    // B column
            sheet.Cells[i, 2].PutValue((i + 1) * 3);    // C column
        }

        // Define the data range for the sparkline
        string dataRange = "A1:C5";

        // Define the location range where sparklines will be placed (E1:E5)
        CellArea location = CellArea.CreateCellArea("E1", "E5");

        // Add a sparkline group of type Line; this also creates sparklines for each row in the location range
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, dataRange, false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Customize the sparkline appearance
        // Set series (line) color to orange
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = Color.Orange;
        group.SeriesColor = seriesColor;

        // Show high and low points with custom colors
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        CellsColor highPointColor = workbook.CreateCellsColor();
        highPointColor.Color = Color.Green;
        group.HighPointColor = highPointColor;

        CellsColor lowPointColor = workbook.CreateCellsColor();
        lowPointColor.Color = Color.Red;
        group.LowPointColor = lowPointColor;

        // Apply a preset style (optional)
        group.PresetStyle = SparklinePresetStyleType.Style5;

        // Save the workbook as an XLSB file
        workbook.Save("SparklineExample.xlsb", SaveFormat.Xlsb);
    }
}