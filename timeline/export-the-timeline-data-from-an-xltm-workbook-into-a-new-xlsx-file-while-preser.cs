using System;
using Aspose.Cells;

class ExportTimeline
{
    static void Main()
    {
        // Load the XLTM workbook that contains the Timeline control
        Workbook sourceWorkbook = new Workbook("TimelineTemplate.xltm");

        // Save the workbook as XLSX. All worksheet content, including the Timeline, is preserved.
        sourceWorkbook.Save("TimelineExported.xlsx", SaveFormat.Xlsx);
    }
}