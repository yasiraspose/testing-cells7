using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsCellManipulation
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook (lifecycle: create) ----------
            Workbook workbook = new Workbook(); // creates an empty XLSX workbook

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // ---------- Manipulate individual cells ----------
            // Put a string value into cell A1 using the string indexer
            worksheet.Cells["A1"].PutValue("Original Value");

            // Retrieve the value from A1
            string a1Value = worksheet.Cells["A1"].StringValue;
            Console.WriteLine($"A1 before change: {a1Value}");

            // Change the value of cell B2 using numeric indexes (row 1, column 1)
            worksheet.Cells[1, 1].PutValue(12345);

            // Retrieve the numeric value from B2
            double b2Value = worksheet.Cells[1, 1].DoubleValue;
            Console.WriteLine($"B2 after change: {b2Value}");

            // Apply a simple style to cell A1 (bold font, yellow background)
            Style style = worksheet.Cells["A1"].GetStyle();
            style.Font.IsBold = true;
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;
            worksheet.Cells["A1"].SetStyle(style);

            // ---------- Save the workbook (lifecycle: save) ----------
            // Save to a file named "ManipulatedCells.xlsx" in the current directory
            workbook.Save("ManipulatedCells.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved as ManipulatedCells.xlsx");
        }
    }
}