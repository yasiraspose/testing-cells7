using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the original XLS timeline spreadsheet
        string sourcePath = "timeline.xls";

        // Desired output path for the converted XLSX file
        string destinationPath = "timeline.xlsx";

        try
        {
            // Convert the XLS file to XLSX format while preserving data and layout
            ConversionUtility.Convert(sourcePath, destinationPath);
            Console.WriteLine($"Conversion successful: '{sourcePath}' → '{destinationPath}'");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}