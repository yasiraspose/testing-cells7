using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class ConvertXlsxToMhtml
{
    static void Main()
    {
        // Input XLSX file and desired MHTML output file
        string sourceFile = "input.xlsx";
        string outputFile = "output.mht";

        // If the source file does not exist, create a simple workbook for demonstration
        if (!System.IO.File.Exists(sourceFile))
        {
            // Create a new workbook (uses the Workbook constructor rule)
            Workbook wb = new Workbook();

            // Add some sample data
            wb.Worksheets[0].Cells["A1"].PutValue("Sample data for conversion");

            // Save the workbook as XLSX (uses the Workbook.Save(string, SaveFormat) rule)
            wb.Save(sourceFile, SaveFormat.Xlsx);
        }

        // Convert the XLSX file to MHTML using the ConversionUtility (conversion rule)
        ConversionUtility.Convert(sourceFile, outputFile);

        Console.WriteLine($"Conversion completed successfully: {sourceFile} -> {outputFile}");
    }
}