using System;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source XLSB file (timeline worksheet)
        string sourcePath = "timeline.xlsb";

        // Desired path for the converted XLSX file
        string destPath = "timeline.xlsx";

        try
        {
            // Convert the XLSB file to XLSX format while preserving data and formatting
            ConversionUtility.Convert(sourcePath, destPath);
            Console.WriteLine($"Conversion successful. Output saved to: {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Conversion failed: {ex.Message}");
        }
    }
}