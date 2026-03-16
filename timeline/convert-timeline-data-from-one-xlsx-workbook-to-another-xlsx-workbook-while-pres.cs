using System;
using Aspose.Cells;

class TimelineTransfer
{
    static void Main()
    {
        // Path to the source workbook containing the timeline
        string sourcePath = "source.xlsx";

        // Path where the destination workbook will be saved
        string destinationPath = "destination.xlsx";

        // Load the source workbook (uses the Workbook(string) constructor)
        Workbook sourceWorkbook = new Workbook(sourcePath);

        // Create an empty workbook for the destination
        Workbook destinationWorkbook = new Workbook();

        // Copy all contents, formatting, and objects (including timelines) from source to destination
        destinationWorkbook.Copy(sourceWorkbook);

        // Save the destination workbook to the specified file
        destinationWorkbook.Save(destinationPath);
    }
}