using System;
using Aspose.Cells;

namespace FormulaLocalLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with actual path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // Scenario 1: Read the localized formula of a cell
            // -----------------------------------------------------------------
            // Assume cell B2 already contains a formula.
            Cell cellB2 = sheet.Cells["B2"];
            Console.WriteLine("Standard Formula (English): " + cellB2.Formula);
            Console.WriteLine("Localized Formula (current locale): " + cellB2.FormulaLocal);

            // -----------------------------------------------------------------
            // Scenario 2: Change workbook locale and observe FormulaLocal change
            // -----------------------------------------------------------------
            // Set workbook region to German (de-DE) – this changes default locale.
            workbook.Settings.Region = CountryCode.Germany;

            // Set a formula using the standard (English) name.
            Cell cellC3 = sheet.Cells["C3"];
            cellC3.Formula = "=SUM(D3:E3)";

            // After changing the region, FormulaLocal reflects the German name.
            Console.WriteLine("\nAfter setting workbook region to Germany:");
            Console.WriteLine("Standard Formula: " + cellC3.Formula);
            Console.WriteLine("Localized Formula: " + cellC3.FormulaLocal); // Expected "SUMME"

            // -----------------------------------------------------------------
            // Scenario 3: Set a formula directly in the local language
            // -----------------------------------------------------------------
            // Use the German function name "SUMME" via FormulaLocal.
            Cell cellD4 = sheet.Cells["D4"];
            cellD4.FormulaLocal = "=SUMME(E4:F4)";

            // Verify that the standard Formula property is automatically translated.
            Console.WriteLine("\nAfter setting FormulaLocal with German function:");
            Console.WriteLine("Standard Formula: " + cellD4.Formula);      // Expected "=SUM(E4:F4)"
            Console.WriteLine("Localized Formula: " + cellD4.FormulaLocal); // Expected "=SUMME(E4:F4)"

            // -----------------------------------------------------------------
            // Scenario 4: Use custom globalization settings for a non‑standard locale
            // -----------------------------------------------------------------
            // Create custom settings that map English "AVERAGE" to French "MOYENNE".
            SettableGlobalizationSettings customSettings = new SettableGlobalizationSettings();
            customSettings.SetLocalFunctionName("AVERAGE", "MOYENNE", true);
            workbook.Settings.GlobalizationSettings = customSettings;

            // Apply a formula using the French name via FormulaLocal.
            Cell cellE5 = sheet.Cells["E5"];
            cellE5.FormulaLocal = "=MOYENNE(F5:G5)";

            // The standard Formula property shows the English equivalent.
            Console.WriteLine("\nUsing custom globalization (French AVERAGE):");
            Console.WriteLine("Standard Formula: " + cellE5.Formula);      // Expected "=AVERAGE(F5:G5)"
            Console.WriteLine("Localized Formula: " + cellE5.FormulaLocal); // Expected "=MOYENNE(F5:G5)"

            // -----------------------------------------------------------------
            // Scenario 5: Parse a locale‑dependent formula string with options
            // -----------------------------------------------------------------
            // When setting a formula that contains locale‑specific separators,
            // enable LocaleDependent in FormulaParseOptions.
            FormulaParseOptions parseOptions = new FormulaParseOptions
            {
                LocaleDependent = true,
                R1C1Style = false
            };

            // Example: French date format with locale identifier.
            Cell cellF6 = sheet.Cells["F6"];
            cellF6.SetFormula("=TEXTE(AUJOURDHUI();\"[$-fr-FR]dddd, dd mmmm yyyy\")", parseOptions);

            Console.WriteLine("\nFormula set with LocaleDependent option:");
            Console.WriteLine("Standard Formula: " + cellF6.Formula);
            Console.WriteLine("Localized Formula: " + cellF6.FormulaLocal);

            // -----------------------------------------------------------------
            // Scenario 6: Retrieve formulas in both formats for export or logging
            // -----------------------------------------------------------------
            for (int row = 1; row <= sheet.Cells.MaxDataRow + 1; row++)
            {
                Cell c = sheet.Cells[row, 1]; // Column A
                if (!string.IsNullOrEmpty(c.Formula))
                {
                    Console.WriteLine($"Cell {c.Name}: English='{c.Formula}' | Local='{c.FormulaLocal}'");
                }
            }

            // Save the modified workbook (replace with desired output path)
            workbook.Save("output.xlsx");
        }
    }
}