using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class ProjectTimelineToNumbers
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample project data: Task, StartDate, EndDate, Owner
        // Header row
        cells[0, 0].Value = "Task";
        cells[0, 1].Value = "StartDate";
        cells[0, 2].Value = "EndDate";
        cells[0, 3].Value = "Owner";

        // Data rows
        cells[1, 0].Value = "Design";
        cells[1, 1].Value = new DateTime(2023, 1, 5);
        cells[1, 2].Value = new DateTime(2023, 1, 20);
        cells[1, 3].Value = "Alice";

        cells[2, 0].Value = "Development";
        cells[2, 1].Value = new DateTime(2023, 1, 21);
        cells[2, 2].Value = new DateTime(2023, 3, 15);
        cells[2, 3].Value = "Bob";

        cells[3, 0].Value = "Testing";
        cells[3, 1].Value = new DateTime(2023, 3, 16);
        cells[3, 2].Value = new DateTime(2023, 4, 5);
        cells[3, 3].Value = "Carol";

        cells[4, 0].Value = "Deployment";
        cells[4, 1].Value = new DateTime(2023, 4, 6);
        cells[4, 2].Value = new DateTime(2023, 4, 10);
        cells[4, 3].Value = "Dave";

        // Create a date style with desired format (e.g., "m/d/yyyy")
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "m/d/yyyy";

        // Apply date style to all date cells
        for (int row = 1; row <= 4; row++)
        {
            cells[row, 1].SetStyle(dateStyle); // StartDate
            cells[row, 2].SetStyle(dateStyle); // EndDate
        }

        // Create a PivotTable that uses the StartDate field (for timeline)
        // Define the source range (A1:D5) and place the pivot table at G1
        int pivotIndex = sheet.PivotTables.Add("A1:D5", "G1", "ProjectPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table
        // Row area: Task
        pivot.AddFieldToArea(PivotFieldType.Row, "Task");
        // Column area: StartDate (will be used for timeline)
        pivot.AddFieldToArea(PivotFieldType.Column, "StartDate");
        // Data area: Owner (just to have a data field)
        pivot.AddFieldToArea(PivotFieldType.Data, "Owner");

        // Refresh and calculate the pivot table
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a Timeline control linked to the pivot table, using the StartDate field
        // Place the timeline at row 12, column 0 (cell A12)
        int timelineIndex = sheet.Timelines.Add(pivot, 12, 0, "StartDate");
        Timeline timeline = sheet.Timelines[timelineIndex];

        // Optional: set a caption for the timeline
        timeline.Caption = "Project Start Date Timeline";

        // Save the workbook as an Apple Numbers file, preserving date ordering and formatting
        workbook.Save("ProjectTimeline.numbers", SaveFormat.Numbers);
    }
}