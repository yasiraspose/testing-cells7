using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CellManipulationDemo
    {
        public static void Run()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the cells collection
            Cells cells = worksheet.Cells;

            // -----------------------------------------------------------------
            // Retrieve cells using numeric indexes (row, column) – zero based
            // -----------------------------------------------------------------
            // Cell at row 0, column 0 corresponds to A1
            Cell cellA1 = cells[0, 0];
            cellA1.PutValue("First Item");               // Set a string value

            // Cell at row 1, column 0 corresponds to A2
            Cell cellA2 = cells[1, 0];
            cellA2.PutValue(12345);                      // Set an integer value

            // -----------------------------------------------------------------
            // Retrieve cells using the A1 style address
            // -----------------------------------------------------------------
            Cell cellB1 = cells["B1"];
            cellB1.PutValue(3.14159);                    // Set a double value

            Cell cellC1 = cells["C1"];
            cellC1.PutValue(DateTime.Now);               // Set a DateTime value

            // -----------------------------------------------------------------
            // Read values back from the cells
            // -----------------------------------------------------------------
            Console.WriteLine("A1 (String): " + cellA1.StringValue);
            Console.WriteLine("A2 (Int): " + cellA2.IntValue);
            Console.WriteLine("B1 (Double): " + cellB1.DoubleValue);
            Console.WriteLine("C1 (DateTime): " + cellC1.DateTimeValue);

            // -----------------------------------------------------------------
            // Modify cell style using GetStyle / SetStyle
            // -----------------------------------------------------------------
            // Make A1 bold and set background color
            Style styleA1 = cellA1.GetStyle();
            styleA1.Font.IsBold = true;
            styleA1.ForegroundColor = Color.Yellow;
            styleA1.Pattern = BackgroundType.Solid;
            cellA1.SetStyle(styleA1);

            // Change B1 font color to red
            Style styleB1 = cellB1.GetStyle();
            styleB1.Font.Color = Color.Red;
            cellB1.SetStyle(styleB1);

            // -----------------------------------------------------------------
            // Save the workbook (lifecycle rule: save)
            // -----------------------------------------------------------------
            workbook.Save("CellManipulationDemo.xlsx", SaveFormat.Xlsx);

            // Dispose the workbook to release resources
            workbook.Dispose();
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CellManipulationDemo.Run();
        }
    }
}