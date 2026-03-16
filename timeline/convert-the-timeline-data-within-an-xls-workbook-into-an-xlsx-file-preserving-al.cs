using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    public class TimelineXlsToXlsxConverter
    {
        public static void Run()
        {
            string sourcePath = "TimelineData.xls";
            string destinationPath = "TimelineData.xlsx";

            try
            {
                Workbook wb = new Workbook(sourcePath);
                wb.Save(destinationPath, SaveFormat.Xlsx);
                Console.WriteLine($"Conversion successful: '{sourcePath}' -> '{destinationPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }

        public static void Main(string[] args)
        {
            Run();
        }
    }
}