using System;
using Aspose.Cells;

namespace AsposeCellsTimelineViaXml
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data required for a timeline (date field)
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["A2"].PutValue(DateTime.Now.AddDays(-3));
            sheet.Cells["A3"].PutValue(DateTime.Now.AddDays(-2));
            sheet.Cells["A4"].PutValue(DateTime.Now.AddDays(-1));
            sheet.Cells["A5"].PutValue(DateTime.Now);

            // Create a pivot table that will serve as the timeline data source
            int pivotIdx = sheet.PivotTables.Add("A1:A5", "C1", "PivotTable1");
            // Note: Adding fields to the pivot table is optional for this example

            // ------------------------------------------------------------
            // Embed SpreadsheetML that defines a Timeline control.
            // ------------------------------------------------------------
            string timelineXml = @"
                <timeline xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""
                          xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
                  <timelineCache pivotCacheId=""1"" />
                  <timelineState>
                    <timelineStateItem name=""date"" />
                  </timelineState>
                </timeline>";

            // Import the XML into the worksheet at cell A1 (row 0, column 0)
            workbook.ImportXml(timelineXml, "Data", 0, 0);

            // Save the workbook with the embedded timeline
            workbook.Save("TimelineViaSpreadsheetML.xlsx");
        }
    }
}