using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class AddSparklineToXlsm
{
    static void Main()
    {
        // Create a new workbook (macro‑enabled compatibility will be set on save)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that the sparkline will represent
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell where the sparkline will be placed (E1)
        CellArea location = CellArea.CreateCellArea("E1", "E1");

        // Add a sparkline group of type Line, using the data range A1:D1,
        // horizontal orientation (isVertical = false), and place it in the defined location
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Optional: customize the appearance of the sparkline group
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = Color.Orange;
        group.SeriesColor = seriesColor;
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        // Save the workbook as an XLSM file to retain macro‑enabled compatibility
        workbook.Save("SparklineDemo.xlsm", SaveFormat.Xlsm);
    }
}