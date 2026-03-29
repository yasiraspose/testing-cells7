using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace SparklineDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that the sparkline will represent
            // Data placed in cells A1 through D1
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the location where the sparkline will be rendered.
            // Here we place it in cell E1 (column index 4, row index 0)
            CellArea sparklineLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group of type Line.
            // Parameters:
            //   SparklineType.Line – type of sparkline
            //   "A1:D1" – data range for the sparkline
            //   false – plot by row (horizontal)
            //   sparklineLocation – where the sparkline cells will be placed
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineLocation);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add a sparkline to the group.
            // The Add method returns the index of the newly created sparkline.
            // Parameters: data range, row index, column index of the sparkline cell.
            int sparkIdx = group.Sparklines.Add("A1:D1", 0, 4);
            Sparkline sparkline = group.Sparklines[sparkIdx];

            // Optional: customize appearance (e.g., show high/low points)
            group.ShowHighPoint = true;
            group.ShowLowPoint = true;
            group.HighPointColor.Color = System.Drawing.Color.Green;
            group.LowPointColor.Color = System.Drawing.Color.Red;
            group.LineWeight = 1.0;

            // Save the workbook containing the sparkline
            workbook.Save("SparklineExample.xls");
        }
    }
}