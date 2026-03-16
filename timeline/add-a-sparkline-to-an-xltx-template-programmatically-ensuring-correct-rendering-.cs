using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using System.Drawing;

class AddSparklineToTemplate
{
    static void Main()
    {
        // Load the XLTX template
        Workbook workbook = new Workbook("Template.xltx");

        // Access the first worksheet (adjust if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data that the sparkline will represent (row 1, columns A‑D)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell where the sparkline will be placed (e.g., E1)
        CellArea sparklineLocation = CellArea.CreateCellArea("E1", "E1");

        // Add a sparkline group:
        // - Type: Line
        // - Data range: A1:D1
        // - Horizontal orientation (isVertical = false)
        // - Location: E1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Optional: customize appearance
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        group.HighPointColor = workbook.CreateCellsColor();
        group.HighPointColor.Color = Color.Green;

        group.LowPointColor = workbook.CreateCellsColor();
        group.LowPointColor.Color = Color.Red;

        group.SeriesColor = workbook.CreateCellsColor();
        group.SeriesColor.Color = Color.Orange;

        group.LineWeight = 1.0;

        // Save the workbook with the added sparkline
        workbook.Save("Result.xlsx", SaveFormat.Xlsx);
    }
}