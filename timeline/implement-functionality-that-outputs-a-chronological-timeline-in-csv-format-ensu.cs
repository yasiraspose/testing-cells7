using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Write header names
        cells[0, 0].PutValue("Date");
        cells[0, 1].PutValue("Value");

        // Sample chronological data
        DateTime[] dates = new DateTime[]
        {
            new DateTime(2021, 1, 1),
            new DateTime(2021, 2, 15),
            new DateTime(2021, 3, 10)
        };
        double[] values = new double[] { 100, 150, 200 };

        // Populate the worksheet with the data
        for (int i = 0; i < dates.Length; i++)
        {
            cells[i + 1, 0].PutValue(dates[i]);
            cells[i + 1, 1].PutValue(values[i]);
        }

        // Apply a date format to the Date column
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "yyyy-MM-dd";
        for (int i = 1; i <= dates.Length; i++)
        {
            cells[i, 0].SetStyle(dateStyle);
        }

        // Configure CSV save options: comma delimiter and keep separators for blank rows
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            Separator = ',',
            KeepSeparatorsForBlankRow = true,
            Encoding = Encoding.UTF8
        };

        // Save the workbook to a memory stream as CSV and output the result
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, saveOptions);
            string csvContent = Encoding.UTF8.GetString(stream.ToArray());
            Console.WriteLine(csvContent);
        }
    }
}