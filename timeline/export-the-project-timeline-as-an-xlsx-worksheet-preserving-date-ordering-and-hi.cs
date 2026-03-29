using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace ProjectTimelineExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Populate worksheet with hierarchical task data
            // Columns: A - Task Name, B - Start Date, C - End Date, D - Level (for hierarchy)
            // ------------------------------------------------------------

            // Header row
            cells["A1"].PutValue("Task");
            cells["B1"].PutValue("Start Date");
            cells["C1"].PutValue("End Date");
            cells["D1"].PutValue("Level");

            // Sample tasks (already ordered by start date)
            // Level 0 = top‑level task, Level 1 = sub‑task, etc.
            cells["A2"].PutValue("Project Initiation");
            cells["B2"].PutValue(new DateTime(2023, 1, 2));
            cells["C2"].PutValue(new DateTime(2023, 1, 5));
            cells["D2"].PutValue(0);

            cells["A3"].PutValue("  Define Scope"); // Indentation for visual hierarchy
            cells["B3"].PutValue(new DateTime(2023, 1, 2));
            cells["C3"].PutValue(new DateTime(2023, 1, 3));
            cells["D3"].PutValue(1);

            cells["A4"].PutValue("  Identify Stakeholders");
            cells["B4"].PutValue(new DateTime(2023, 1, 3));
            cells["C4"].PutValue(new DateTime(2023, 1, 4));
            cells["D4"].PutValue(1);

            cells["A5"].PutValue("Project Planning");
            cells["B5"].PutValue(new DateTime(2023, 1, 6));
            cells["C5"].PutValue(new DateTime(2023, 1, 15));
            cells["D5"].PutValue(0);

            cells["A6"].PutValue("  Create Schedule");
            cells["B6"].PutValue(new DateTime(2023, 1, 6));
            cells["C6"].PutValue(new DateTime(2023, 1, 8));
            cells["D6"].PutValue(1);

            cells["A7"].PutValue("  Resource Allocation");
            cells["B7"].PutValue(new DateTime(2023, 1, 9));
            cells["C7"].PutValue(new DateTime(2023, 1, 11));
            cells["D7"].PutValue(1);

            cells["A8"].PutValue("Project Execution");
            cells["B8"].PutValue(new DateTime(2023, 1, 16));
            cells["C8"].PutValue(new DateTime(2023, 2, 28));
            cells["D8"].PutValue(0);

            // ------------------------------------------------------------
            // Apply date style to date columns (B and C)
            // ------------------------------------------------------------
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "yyyy-mm-dd";

            for (int row = 2; row <= 8; row++)
            {
                cells[row - 1, 1].SetStyle(dateStyle); // Column B (Start Date)
                cells[row - 1, 2].SetStyle(dateStyle); // Column C (End Date)
            }

            // ------------------------------------------------------------
            // Group rows to reflect hierarchical structure
            // ------------------------------------------------------------
            // Group rows 3‑4 under row 2 (Project Initiation)
            sheet.Cells.GroupRows(2, 3, true);
            // Group rows 6‑7 under row 5 (Project Planning)
            sheet.Cells.GroupRows(5, 6, true);

            // ------------------------------------------------------------
            // (Optional) Add a PivotTable and Timeline for interactive view
            // ------------------------------------------------------------
            // Define data range for the pivot table (including header)
            string dataRange = "A1:D8";

            // Add pivot table at cell G2
            int pivotIndex = sheet.PivotTables.Add(dataRange, "G2", "TimelinePivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Use Task as row field, Start Date as column field, and count of tasks as data field
            pivot.AddFieldToArea(PivotFieldType.Row, "Task");
            pivot.AddFieldToArea(PivotFieldType.Column, "Start Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Task");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

            // Refresh and calculate pivot data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline control linked to the pivot table (based on Start Date)
            // Place the timeline at cell G15
            sheet.Timelines.Add(pivot, "G15", "Start Date");

            // ------------------------------------------------------------
            // Save the workbook as XLSX (lifecycle: save)
            // ------------------------------------------------------------
            workbook.Save("ProjectTimeline.xlsx");
        }
    }
}