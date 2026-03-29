using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class ProjectTimelineTemplate
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Populate worksheet with sample project data
        // Columns: Task | StartDate | EndDate
        // -------------------------------------------------
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["B1"].PutValue("StartDate");
        sheet.Cells["C1"].PutValue("EndDate");

        sheet.Cells["A2"].PutValue("Planning");
        sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 1));
        sheet.Cells["C2"].PutValue(new DateTime(2023, 1, 15));

        sheet.Cells["A3"].PutValue("Design");
        sheet.Cells["B3"].PutValue(new DateTime(2023, 1, 16));
        sheet.Cells["C3"].PutValue(new DateTime(2023, 2, 5));

        sheet.Cells["A4"].PutValue("Development");
        sheet.Cells["B4"].PutValue(new DateTime(2023, 2, 6));
        sheet.Cells["C4"].PutValue(new DateTime(2023, 4, 30));

        sheet.Cells["A5"].PutValue("Testing");
        sheet.Cells["B5"].PutValue(new DateTime(2023, 5, 1));
        sheet.Cells["C5"].PutValue(new DateTime(2023, 5, 31));

        // -------------------------------------------------
        // Create a PivotTable that will serve as the data source for the Timeline
        // Use the StartDate field as the date axis
        // -------------------------------------------------
        int pivotIndex = sheet.PivotTables.Add("A1:C5", "E1", "ProjectPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Row field: StartDate (date axis)
        pivot.AddFieldToArea(PivotFieldType.Row, "StartDate");
        // Column field: Task (to differentiate tasks)
        pivot.AddFieldToArea(PivotFieldType.Column, "Task");
        // Data field: EndDate (we just need a numeric aggregation; Count works)
        pivot.AddFieldToArea(PivotFieldType.Data, "EndDate");

        // Refresh and calculate the pivot table
        pivot.RefreshData();
        pivot.CalculateData();

        // -------------------------------------------------
        // Add a Timeline control linked to the PivotTable
        // Position it starting at cell A10
        // -------------------------------------------------
        int timelineIndex = sheet.Timelines.Add(pivot, "A10", "StartDate");
        Timeline timeline = sheet.Timelines[timelineIndex];

        // Optional: customize timeline appearance
        timeline.Caption = "Project Timeline";
        timeline.ShowHeader = true;
        timeline.ShowHorizontalScrollbar = true;
        timeline.ShowSelectionLabel = true;
        timeline.ShowTimeLevel = true;

        // -------------------------------------------------
        // Save the workbook as an Excel template (XLTX)
        // -------------------------------------------------
        workbook.Save("ProjectTimelineTemplate.xltx", SaveFormat.Xltx);
    }
}