using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing XLSX workbook (create rule)
            Workbook workbook = new Workbook("input.xlsx");

            // Set the workbook's locale to German (example of localization)
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and cell A1
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            // Set a formula using the standard (English) syntax
            cell.Formula = "=SUM(B1:C1)";

            // Display the formula in both standard and localized forms
            Console.WriteLine("After setting standard formula:");
            Console.WriteLine("Standard Formula   : " + cell.Formula);
            Console.WriteLine("Localized Formula  : " + cell.FormulaLocal);

            // Set a formula using the localized (German) syntax
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // Display the formulas again to show the conversion
            Console.WriteLine("\nAfter setting localized formula:");
            Console.WriteLine("Standard Formula   : " + cell.Formula);
            Console.WriteLine("Localized Formula  : " + cell.FormulaLocal);

            // Calculate all formulas to update cell values
            workbook.CalculateFormula();

            // Show the calculated value of the cell
            Console.WriteLine("\nCalculated Value   : " + cell.Value);

            // Save the modified workbook (save rule)
            workbook.Save("output.xlsx");
        }
    }
}