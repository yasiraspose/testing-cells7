using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the XLTX template that contains the timeline
        string templatePath = "Template.xltx";

        // Load the template workbook (preserves all objects, including timelines)
        Workbook workbook = new Workbook(templatePath);

        // Verify that timelines are present (optional, for debugging)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            var timelines = sheet.Timelines;
            Console.WriteLine($"Worksheet '{sheet.Name}' contains {timelines.Count} timeline(s).");
        }

        // Save the workbook as a regular XLSX file, retaining all content
        string outputPath = "Result.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}