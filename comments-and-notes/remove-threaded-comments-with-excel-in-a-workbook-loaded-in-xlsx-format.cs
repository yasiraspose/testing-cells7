using System;
using Aspose.Cells;

class RemoveThreadedCommentsDemo
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range to limit the search for threaded comments
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Scan each cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    // Retrieve threaded comments for the current cell
                    ThreadedCommentCollection threadedComments = sheet.Comments.GetThreadedComments(row, col);

                    // If any threaded comments exist, clear them
                    if (threadedComments != null && threadedComments.Count > 0)
                    {
                        threadedComments.Clear(); // Remove all threaded comments from this cell
                    }
                }
            }
        }

        // Save the workbook after removing all threaded comments
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}