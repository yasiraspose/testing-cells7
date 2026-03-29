using System;
using Aspose.Cells;

namespace TimelineConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the macro‑enabled workbook (XLSM)
            string sourcePath = "timeline.xlsm";

            // Desired path for the macro‑free workbook (XLSX)
            string destinationPath = "timeline.xlsx";

            // Load the existing XLSM file
            Workbook workbook = new Workbook(sourcePath);

            // Remove any VBA/macros from the workbook
            workbook.RemoveMacro();

            // Save the cleaned workbook as a standard XLSX file
            workbook.Save(destinationPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed: '{sourcePath}' → '{destinationPath}'");
        }
    }
}