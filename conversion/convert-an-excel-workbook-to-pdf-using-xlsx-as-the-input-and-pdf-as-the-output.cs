using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "source.xlsx";

            // Desired output PDF file path
            string destPath = "output.pdf";

            // Load the Excel workbook from the file (create & load lifecycle)
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as PDF (save lifecycle)
            workbook.Save(destPath, SaveFormat.Pdf);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}