using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class EmbedTimelineFromMhtml
{
    static void Main()
    {
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.MHtml);
        Workbook workbook = new Workbook("input.mht", loadOptions);

        Worksheet sheet = workbook.Worksheets[0];

        int pivotIndex = sheet.PivotTables.Add("A1:C5", "E1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        pivot.AddFieldToArea(PivotFieldType.Row, 0);
        pivot.AddFieldToArea(PivotFieldType.Column, 1);
        pivot.AddFieldToArea(PivotFieldType.Data, 2);

        pivot.RefreshData();
        pivot.CalculateData();

        sheet.Timelines.Add(pivot, "H1", "Date");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);
        saveOptions.PresentationPreference = true;
        workbook.Save("output.mht", saveOptions);
    }
}