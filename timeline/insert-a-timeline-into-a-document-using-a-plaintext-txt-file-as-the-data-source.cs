using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineFromTxt
{
    public class Program
    {
        public static void Main()
        {
            // Path to the plain‑text data source (tab‑ or comma‑separated values)
            string txtPath = "data.txt";

            // Read all lines from the TXT file
            string[] lines = File.ReadAllLines(txtPath);

            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate worksheet with data from the TXT file
            for (int row = 0; row < lines.Length; row++)
            {
                // Split line by comma (you can change to '\t' for tab‑separated)
                string[] parts = lines[row].Split(',');

                for (int col = 0; col < parts.Length; col++)
                {
                    // Try to parse dates and numbers, otherwise treat as string
                    string value = parts[col].Trim();

                    if (DateTime.TryParse(value, out DateTime dt))
                        cells[row, col].Value = dt;
                    else if (double.TryParse(value, out double num))
                        cells[row, col].Value = num;
                    else
                        cells[row, col].Value = value;
                }
            }

            // Assume the first column is a date field named "Date"
            // and the second column is a numeric field named "Value"
            // Adjust field names according to your TXT header line
            string dateFieldName = cells[0, 0].StringValue;   // e.g., "Date"
            string valueFieldName = cells[0, 1].StringValue; // e.g., "Value"

            // Determine the data range (excluding header row)
            int totalRows = lines.Length;
            int totalCols = lines[0].Split(',').Length;
            string dataRange = $"A1:{CellIndexToName(totalRows - 1, totalCols - 1)}";

            // Add a PivotTable using the data range (pivot table creation rule)
            int pivotIndex = sheet.PivotTables.Add(dataRange, "E3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add the date field to the Row area and the value field to the Data area
            pivot.AddFieldToArea(PivotFieldType.Row, dateFieldName);
            pivot.AddFieldToArea(PivotFieldType.Data, valueFieldName);

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the pivot table (timeline addition rule)
            // Place the timeline starting at cell G1 and bind it to the date field
            sheet.Timelines.Add(pivot, "G1", dateFieldName);

            // Save the workbook (save rule)
            workbook.Save("TimelineFromTxt.xlsx");
        }

        // Helper method to convert zero‑based row/column indexes to Excel cell name (e.g., 0,0 -> A1)
        private static string CellIndexToName(int row, int col)
        {
            // Convert column index to letters
            string colName = "";
            int dividend = col + 1;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                colName = Convert.ToChar('A' + modulo) + colName;
                dividend = (dividend - modulo) / 26;
            }

            // Row index is zero‑based, Excel rows start at 1
            return $"{colName}{row + 1}";
        }
    }
}