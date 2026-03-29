using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineFromXml
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Import XML data into the first worksheet starting at cell A1
            // The XML file should contain at least a Date column and a numeric column (e.g., Value)
            workbook.ImportXml("timelineDefinition.xml", "Sheet1", 0, 0);

            // Get reference to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range after importing XML
            int lastRow = sheet.Cells.MaxDataRow; // zero‑based index
            // Assuming the XML data has two columns: Date (A) and Value (B)
            string dataRange = $"A1:B{lastRow + 1}";

            // Add a PivotTable based on the imported data
            int pivotIndex = sheet.PivotTables.Add(dataRange, "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the PivotTable: Date as row field, Value as data field
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh and calculate the PivotTable data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the PivotTable, positioned at cell F1, using the "Date" field
            int timelineIndex = sheet.Timelines.Add(pivot, "F1", "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Optional: set some visual properties of the Timeline
            timeline.Caption = "Sales Timeline";
            timeline.Name = "SalesTimeline";

            // Save the workbook with the embedded Timeline
            workbook.Save("TimelineFromXml.xlsx");
        }
    }
}