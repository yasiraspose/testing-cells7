using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (XLSX)
            string sourcePath = "input.xlsx";

            // Desired output path for the XPS file
            string outputPath = "output.xps";

            // Convert the Excel workbook to XPS using Aspose.Cells ConversionUtility
            // This utilizes the provided Convert(string, string) method.
            ConversionUtility.Convert(sourcePath, outputPath);

            Console.WriteLine("Conversion from XLSX to XPS completed successfully.");
        }
    }
}