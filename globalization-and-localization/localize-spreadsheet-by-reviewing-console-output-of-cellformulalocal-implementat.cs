using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string inputPath = "sample.xlsx";

            // Load the workbook (uses the standard load lifecycle)
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook region to German to obtain German localized formulas
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Iterate through all cells that contain data
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that have a formula
                    if (cell.IsFormula)
                    {
                        // Standard (English) formula
                        string standardFormula = cell.Formula;

                        // Localized formula obtained via the FormulaLocal property
                        string localizedFormula = cell.FormulaLocal;

                        // Localized formula obtained via GetFormula (isLocal = true)
                        string localizedViaGet = cell.GetFormula(false, true);

                        // Output the results to the console
                        Console.WriteLine($"Cell {cell.Name}:");
                        Console.WriteLine($"  Standard Formula : {standardFormula}");
                        Console.WriteLine($"  FormulaLocal     : {localizedFormula}");
                        Console.WriteLine($"  GetFormula(true) : {localizedViaGet}");
                        Console.WriteLine();
                    }
                }
            }

            // Save the workbook (optional, demonstrates the save lifecycle)
            workbook.Save("localized_output.xlsx");
        }
    }
}