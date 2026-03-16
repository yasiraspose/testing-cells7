using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSB workbook containing the timeline
        string sourcePath = "timeline.xlsb";

        // Desired path for the converted XLSX workbook
        string destPath = "timeline_converted.xlsx";

        // Load the XLSB workbook
        var loadOptions = new LoadOptions(LoadFormat.Xlsb);
        var workbook = new Workbook(sourcePath, loadOptions);

        // Save as XLSX, preserving all data structures (including timelines)
        workbook.Save(destPath, SaveFormat.Xlsx);

        Console.WriteLine($"Conversion completed successfully: {sourcePath} -> {destPath}");
    }
}