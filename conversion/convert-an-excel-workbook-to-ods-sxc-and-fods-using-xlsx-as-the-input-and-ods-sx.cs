using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConvertXlsxToMultipleFormats
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string odsPath = "output.ods";
            string sxcPath = "output.sxc";
            string fodsPath = "output.fods";

            Workbook workbook = new Workbook(sourcePath);
            workbook.Save(odsPath, SaveFormat.Ods);
            workbook.Save(sxcPath, SaveFormat.Sxc);
            workbook.Save(fodsPath, SaveFormat.Fods);

            Console.WriteLine("Conversion to ODS, SXC, and FODS completed successfully.");
        }
    }
}