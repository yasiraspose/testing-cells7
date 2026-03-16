using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace TimelineTsvExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate worksheet with sample data (Date and Value columns)
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue(new DateTime(2023, 1, 1));
            cells["B2"].PutValue(100);
            cells["A3"].PutValue(new DateTime(2023, 1, 2));
            cells["B3"].PutValue(150);
            cells["A4"].PutValue(new DateTime(2023, 1, 3));
            cells["B4"].PutValue(200);

            // Add a pivot table that uses the Date field as a row field
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table (placed at cell E1)
            // Using the overload: Add(PivotTable, string destCellName, string baseFieldName)
            sheet.Timelines.Add(pivot, "E1", "Date");

            // Configure TSV save options (tab separator)
            TxtSaveOptions tsvOptions = new TxtSaveOptions(SaveFormat.Tsv);
            tsvOptions.Separator = '\t'; // Tab character

            // Save the workbook as a TSV file
            workbook.Save("TimelineData.tsv", tsvOptions);
        }
    }
}