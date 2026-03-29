using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsPrintPrnDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a workbook and populate it with sample data
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet sheet = workbook.Worksheets[0];              // get first worksheet
            Cells cells = sheet.Cells;

            // Sample data: dates and values
            cells["A1"].Value = "Date";
            cells["B1"].Value = "Value";

            cells["A2"].Value = new DateTime(2023, 1, 1);
            cells["B2"].Value = 100;

            cells["A3"].Value = new DateTime(2023, 2, 1);
            cells["B3"].Value = 150;

            cells["A4"].Value = new DateTime(2023, 3, 1);
            cells["B4"].Value = 200;

            // ------------------------------------------------------------
            // 2. Create a PivotTable that will serve as the data source for the Timeline
            // ------------------------------------------------------------
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // ------------------------------------------------------------
            // 3. Add a Timeline linked to the PivotTable
            // ------------------------------------------------------------
            // Place the Timeline starting at cell F1
            int timelineIdx = sheet.Timelines.Add(pivot, "F1", "Date");
            Timeline timeline = sheet.Timelines[timelineIdx];
            timeline.Caption = "Sales Timeline";

            // ------------------------------------------------------------
            // 4. Configure print options to preserve layout and pagination
            // ------------------------------------------------------------
            ImageOrPrintOptions printOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true,               // each page prints on a separate sheet
                PrintWithStatusDialog = false,        // suppress status dialog
                // The following properties can be adjusted as needed:
                // HorizontalResolution = 300,
                // VerticalResolution = 300,
            };

            // Optional: define a print area that includes the timeline
            // This ensures only the relevant range is printed.
            sheet.PageSetup.PrintArea = "A1:B4,F1:G10";

            // ------------------------------------------------------------
            // 5. Create a SheetRender instance for the worksheet
            // ------------------------------------------------------------
            SheetRender sheetRender = new SheetRender(sheet, printOptions);

            // ------------------------------------------------------------
            // 6. Send the rendered pages to a printer that outputs PRN files.
            //    Replace "YourPrinterName" with the actual printer name that
            //    generates PRN output (e.g., a generic text printer driver).
            // ------------------------------------------------------------
            try
            {
                // Use dynamic invocation to call the overload that accepts only the printer name.
                // This avoids compile‑time binding issues if the overload is not present in the
                // referenced version of Aspose.Cells.
                dynamic dynRender = sheetRender;
                dynRender.ToPrinter("YourPrinterName");   // printer must be configured to produce PRN
                Console.WriteLine("Print job sent to printer successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during printing: {ex.Message}");
            }
            finally
            {
                // Release resources used by the renderer
                sheetRender.Dispose();
            }

            // ------------------------------------------------------------
            // 7. Save the workbook for reference (optional)
            // ------------------------------------------------------------
            workbook.Save("TimelinePrintDemo.xlsx");
        }
    }
}