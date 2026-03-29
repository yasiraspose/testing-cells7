using System;
using Aspose.Cells;

namespace TimelineExportExample
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook that contains the Timeline control
            Workbook sourceWorkbook = new Workbook("SourceWorkbook.xlsx");

            // Create a new empty workbook that will receive the exported content
            Workbook destinationWorkbook = new Workbook();

            // Copy all worksheets, including the Timeline, cell data and formatting, from the source
            destinationWorkbook.Copy(sourceWorkbook);

            // Save the new workbook; the Timeline control and its associated data are preserved
            destinationWorkbook.Save("ExportedTimeline.xlsx", SaveFormat.Xlsx);
        }
    }
}