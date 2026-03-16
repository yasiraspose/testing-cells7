using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace TimelineConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSM workbook that contains the Timeline
            string sourcePath = "TimelineWorkbook.xlsm";

            // Desired output path for the converted XLSX workbook
            string destinationPath = "TimelineWorkbook.xlsx";

            try
            {
                // Convert the XLSM file to XLSX format.
                // This uses the ConversionUtility.Convert method which preserves all data
                // (including the Timeline) while changing the file format.
                ConversionUtility.Convert(sourcePath, destinationPath);

                Console.WriteLine($"Conversion successful: '{sourcePath}' -> '{destinationPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}