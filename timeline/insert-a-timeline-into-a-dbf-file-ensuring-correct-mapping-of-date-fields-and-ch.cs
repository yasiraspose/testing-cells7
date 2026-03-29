using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Saving;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // -------------------------------------------------
        // 1. Populate worksheet with sample data (chronological order)
        // -------------------------------------------------
        // Header row
        cells["A1"].PutValue("Date");
        cells["B1"].PutValue("Sales");

        // Sample data – dates are added in chronological order
        cells["A2"].PutValue(new DateTime(2023, 1, 1));
        cells["B2"].PutValue(1500);

        cells["A3"].PutValue(new DateTime(2023, 1, 5));
        cells["B3"].PutValue(2300);

        cells["A4"].PutValue(new DateTime(2023, 1, 10));
        cells["B4"].PutValue(3200);

        cells["A5"].PutValue(new DateTime(2023, 1, 15));
        cells["B5"].PutValue(4100);

        cells["A6"].PutValue(new DateTime(2023, 1, 20));
        cells["B6"].PutValue(5000);

        // -------------------------------------------------
        // 2. Create a PivotTable that uses the Date field as a row field
        // -------------------------------------------------
        // Add the pivot table; source range includes header and data
        int pivotIndex = sheet.PivotTables.Add("A1:B6", "D1", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add Date to the Row area and Sales to the Data area
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Refresh and calculate the pivot table to ensure it reflects the data
        pivot.RefreshData();
        pivot.CalculateData();

        // -------------------------------------------------
        // 3. Add a Timeline linked to the PivotTable's Date field
        // -------------------------------------------------
        // The timeline will be placed starting at cell E1
        sheet.Timelines.Add(pivot, "E1", "Date");

        // -------------------------------------------------
        // 4. Save the workbook as a DBF file with appropriate options
        // -------------------------------------------------
        DbfSaveOptions dbfOptions = new DbfSaveOptions
        {
            // Export dates as strings to preserve formatting in DBF
            ExportAsString = true
        };

        // Save the workbook; the timeline itself is an Excel UI feature,
        // but the underlying data (including the date field) will be stored in the DBF.
        workbook.Save("TimelineData.dbf", dbfOptions);
    }
}