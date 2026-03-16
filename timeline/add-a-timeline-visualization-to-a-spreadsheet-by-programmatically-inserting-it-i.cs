using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace TimelineSxcDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with a date field (required for timeline)
            cells["A1"].Value = "Date";
            cells["B1"].Value = "Sales";

            cells["A2"].Value = new DateTime(2023, 1, 1);
            cells["B2"].Value = 1500;

            cells["A3"].Value = new DateTime(2023, 2, 1);
            cells["B3"].Value = 2000;

            cells["A4"].Value = new DateTime(2023, 3, 1);
            cells["B4"].Value = 1800;

            // Add a pivot table that will serve as the data source for the timeline
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            // Insert a timeline linked to the pivot table.
            // The timeline will be placed with its upper‑left corner at cell E1
            // and will use the "Date" field as the base field.
            sheet.Timelines.Add(pivot, "E1", "Date");

            // Save the workbook in SXC (OpenOffice Spreadsheet) format
            workbook.Save("TimelineDemo.sxc", SaveFormat.Sxc);
        }
    }
}