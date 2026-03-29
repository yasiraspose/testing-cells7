using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the XLTX template that contains the Timeline
        string templatePath = "TimelineTemplate.xltx";

        // Path where the converted XLSX workbook will be saved
        string outputPath = "TimelineWorkbook.xlsx";

        // Load the template workbook
        Workbook workbook = new Workbook(templatePath);

        // Optional: verify that a Timeline exists in the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        if (sheet.Timelines.Count > 0)
        {
            Console.WriteLine($"Timeline count in sheet '{sheet.Name}': {sheet.Timelines.Count}");
        }
        else
        {
            Console.WriteLine("No Timeline found in the loaded template.");
        }

        // Save the workbook as a regular XLSX file, preserving all data including the Timeline
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Template converted and saved to '{outputPath}'.");
    }
}