using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsMaxDisplayRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Load the workbook from an existing XLSX file
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the maximal display range which includes data, merged cells and shapes
            AsposeRange maxDisplayRange = worksheet.Cells.MaxDisplayRange;

            // If the worksheet is empty, MaxDisplayRange will be null (Aspose.Cells 21.5.2+)
            if (maxDisplayRange == null)
            {
                Console.WriteLine("The worksheet is empty. No display range found.");
                return;
            }

            // Calculate the last row and column indices based on the range dimensions
            int firstRow = maxDisplayRange.FirstRow;               // zero‑based index
            int firstColumn = maxDisplayRange.FirstColumn;         // zero‑based index
            int rowCount = maxDisplayRange.RowCount;
            int columnCount = maxDisplayRange.ColumnCount;
            int lastRow = firstRow + rowCount - 1;
            int lastColumn = firstColumn + columnCount - 1;

            // Output the range boundaries
            Console.WriteLine("Max Display Range:");
            Console.WriteLine($"First Row (0‑based): {firstRow}");
            Console.WriteLine($"First Column (0‑based): {firstColumn}");
            Console.WriteLine($"Last Row (0‑based): {lastRow}");
            Console.WriteLine($"Last Column (0‑based): {lastColumn}");
            Console.WriteLine($"Total Rows: {rowCount}");
            Console.WriteLine($"Total Columns: {columnCount}");

            // Optionally, save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
    }
}