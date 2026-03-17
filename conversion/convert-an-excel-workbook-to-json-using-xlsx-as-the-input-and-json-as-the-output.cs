using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Path where the JSON output will be saved
        string jsonPath = "output.json";

        // Load the Excel workbook from the file
        Workbook workbook = new Workbook(sourcePath);

        // Configure JSON save options (optional customizations)
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export as a JSON object even if there is only one worksheet
            AlwaysExportAsJsonObject = true,
            // Treat the first row as header (if present)
            HasHeaderRow = true,
            // Export empty cells as null values
            ExportEmptyCells = true,
            // Use indentation for better readability
            Indent = "  "
        };

        // Save the workbook as a JSON file using the configured options
        workbook.Save(jsonPath, jsonOptions);

        Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{jsonPath}'");
    }
}