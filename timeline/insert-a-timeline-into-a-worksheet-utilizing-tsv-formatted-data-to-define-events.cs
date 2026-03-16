using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineFromTsv
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // TSV data defining events and their timestamps
        string tsvData = @"Event	Date
Launch	2023-01-01
Update	2023-02-15
Retire	2023-12-31";

        // Split the TSV into lines
        string[] lines = tsvData.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

        // Import each line into the worksheet (starting at row 0, column 0)
        for (int i = 0; i < lines.Length; i++)
        {
            string[] parts = lines[i].Split('\t');
            // Parse the date string into a DateTime object
            DateTime dateValue = DateTime.Parse(parts[1]);

            // Prepare an object array for the current row: Event (string) and Date (DateTime)
            object[] rowData = new object[] { parts[0], dateValue };

            // Import the row horizontally (isVertical = false)
            cells.ImportObjectArray(rowData, i, 0, false);
        }

        // Define the data range for the pivot table (A1:B4 in this example)
        // Adjust the range if the TSV has a different number of rows
        string dataRange = "A1:B4";

        // Add a pivot table that will serve as the data source for the timeline
        int pivotIndex = sheet.PivotTables.Add(dataRange, "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Use the Date field as the row field (required for timeline)
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");

        // Refresh and calculate the pivot table data
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a timeline linked to the pivot table, placing it starting at cell F1
        // Use the overload Add(PivotTable, string, string)
        int timelineIndex = sheet.Timelines.Add(pivot, "F1", "Date");

        // Optional: set a name for the timeline
        Timeline timeline = sheet.Timelines[timelineIndex];
        timeline.Name = "EventTimeline";

        // Save the workbook
        workbook.Save("TimelineFromTsv.xlsx");
    }
}