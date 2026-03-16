using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Saving;

namespace TimelinePdfDemo
{
    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Populate worksheet with chronological data (Date, Event, Value)
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Event");
            sheet.Cells["C1"].PutValue("Value");

            DateTime start = new DateTime(2023, 1, 1);
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(start.AddDays(i * 30));
                sheet.Cells[i + 1, 1].PutValue($"Milestone {i + 1}");
                sheet.Cells[i + 1, 2].PutValue((i + 1) * 100);
            }

            // Apply date format
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "yyyy-mm-dd";
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            sheet.Cells.CreateRange("A2:A11").ApplyStyle(dateStyle, flag);

            // 3. Create a PivotTable
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("A1:C11", "E3", "PivotTable1");
            PivotTable pivot = pivots[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Column, "Event");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // 4. Add a Timeline linked to the PivotTable's Date field
            int timelineIndex = sheet.Timelines.Add(pivot, 14, 0, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];
            timeline.Caption = "Project Timeline";
            timeline.ShowHeader = true;
            timeline.ShowHorizontalScrollbar = true;
            timeline.ShowSelectionLabel = true;
            timeline.ShowTimeLevel = true;

            // 5. Adjust page setup
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.PrintArea = "A1:Z50";
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 1;

            // 6. Save as PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true,
                CreatedTime = DateTime.Now
            };
            workbook.Save("ChronologicalTimeline.pdf", pdfOptions);

            Console.WriteLine("PDF with chronological timeline generated successfully.");
        }
    }
}