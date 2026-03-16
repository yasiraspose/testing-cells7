using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        string templatePath = "template.xltm";
        string outputPath = "converted.xlsx";

        Workbook workbook = new Workbook(templatePath);

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            var timelines = sheet.Timelines;
            // Optional: iterate timelines if needed
            // for (int i = 0; i < timelines.Count; i++) { Timeline tl = timelines[i]; }
        }

        workbook.Save(outputPath, SaveFormat.Xlsx);
        Console.WriteLine("XLTM template successfully converted to XLSX.");
    }
}