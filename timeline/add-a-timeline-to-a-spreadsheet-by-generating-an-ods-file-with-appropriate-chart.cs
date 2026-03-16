using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Charts;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data with a Date column, a Category column and a Value column
        cells["A1"].PutValue("Date");
        cells["B1"].PutValue("Category");
        cells["C1"].PutValue("Value");

        DateTime baseDate = new DateTime(2023, 1, 1);
        for (int i = 0; i < 10; i++)
        {
            cells[1 + i, 0].PutValue(baseDate.AddDays(i));          // Date
            cells[1 + i, 1].PutValue("Item " + (i + 1));           // Category
            cells[1 + i, 2].PutValue((i + 1) * 10);                // Value
        }

        // Create a pivot table based on the data range
        PivotTableCollection pivots = sheet.PivotTables;
        int pivotIndex = pivots.Add("A1:C11", "E3", "PivotTable1");
        PivotTable pivot = pivots[pivotIndex];

        // Configure pivot fields: Date as row, Category as column, Value as data
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Column, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a timeline linked to the pivot table, positioned at cell G1, using the Date field
        sheet.Timelines.Add(pivot, "G1", "Date");

        // Add a simple column chart for the raw data (optional visual aid)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 7);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("=Sheet1!$A$1:$C$11", true);
        chart.Title.Text = "Sample Data Chart";

        // Save the workbook as an ODS file using the LibreOffice generator
        OdsSaveOptions saveOptions = new OdsSaveOptions
        {
            GeneratorType = OdsGeneratorType.LibreOffice
        };
        workbook.Save("TimelineDemo.ods", saveOptions);
    }
}