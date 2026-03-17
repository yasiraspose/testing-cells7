using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    public class ExcelToHtmlConverter
    {
        public static void Run()
        {
            // Path to the source Excel file (XLSX)
            string sourcePath = "input.xlsx";

            // Path for the resulting HTML file
            string outputPath = "output.html";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Optional: set the HTML version to HTML5
            htmlOptions.HtmlVersion = HtmlVersion.Html5;

            // Save the workbook as HTML using the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{outputPath}'");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExcelToHtmlConverter.Run();
        }
    }
}