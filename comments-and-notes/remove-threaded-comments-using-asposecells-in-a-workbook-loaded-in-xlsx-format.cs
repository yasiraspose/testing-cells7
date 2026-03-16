using System;
using Aspose.Cells;

class RemoveThreadedCommentsDemo
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet and clear all comments (including threaded comments)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // This method removes both regular and threaded comments from the worksheet
            sheet.ClearComments();
        }

        // Save the workbook after removing the comments
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}