using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class AddSparklineToXltmTemplate
{
    static void Main()
    {
        // Path to the existing macro‑enabled template (XLTM)
        string templatePath = "Template.xltm";

        // Load the template workbook
        Workbook workbook = new Workbook(templatePath);

        // Access the first worksheet (modify as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that the sparkline will represent
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell where the sparkline will be placed (E1 in this case)
        CellArea sparklineLocation = new CellArea
        {
            StartRow = 0,   // Row 1 (zero‑based)
            EndRow = 0,
            StartColumn = 4, // Column E (zero‑based)
            EndColumn = 4
        };

        // Add a sparkline group of type Line, linking it to the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Line,
            sheet.Name + "!A1:D1",
            false,               // plot by column (horizontal)
            sparklineLocation);

        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Optional: customize the sparkline appearance
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;
        group.LineWeight = 1.0;

        // Save the workbook back as a macro‑enabled template
        string outputPath = "Output.xltm";
        workbook.Save(outputPath, SaveFormat.Xltm);
    }
}