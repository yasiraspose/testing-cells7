using System;
using System.IO;
using System.Xml;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineXmlExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate worksheet with sample data (fruit, date, amount)
            cells[0, 0].Value = "Fruit";
            cells[0, 1].Value = "Date";
            cells[0, 2].Value = "Amount";

            cells[1, 0].Value = "Grape";
            cells[1, 1].Value = new DateTime(2021, 2, 5);
            cells[1, 2].Value = 50;

            cells[2, 0].Value = "Blueberry";
            cells[2, 1].Value = new DateTime(2022, 3, 8);
            cells[2, 2].Value = 60;

            cells[3, 0].Value = "Kiwi";
            cells[3, 1].Value = new DateTime(2023, 4, 10);
            cells[3, 2].Value = 70;

            cells[4, 0].Value = "Cherry";
            cells[4, 1].Value = new DateTime(2024, 5, 16);
            cells[4, 2].Value = 80;

            // Create a PivotTable based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("=Sheet1!A1:C5", "E1", "FruitPivot");
            PivotTable pivot = pivots[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Column, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the PivotTable (using the Date field)
            int timelineIndex = sheet.Timelines.Add(pivot, 10, 5, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];
            timeline.Caption = "Fruit Sales Timeline";

            // OPTIONAL: Save the workbook (demonstrates lifecycle usage)
            workbook.Save("TimelineDemo.xlsx");

            // Generate XML representation of the chronological events
            // The XML will contain each date and its associated amount
            string xmlPath = "TimelineData.xml";
            using (XmlWriter writer = XmlWriter.Create(xmlPath, new XmlWriterSettings { Indent = true }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Timeline");
                writer.WriteAttributeString("Caption", timeline.Caption ?? string.Empty);

                // Iterate through the data rows (skip header row)
                for (int row = 1; row <= 4; row++)
                {
                    DateTime date = cells[row, 1].DateTimeValue;
                    double amount = cells[row, 2].DoubleValue;

                    writer.WriteStartElement("Event");
                    writer.WriteElementString("Date", date.ToString("yyyy-MM-dd"));
                    writer.WriteElementString("Amount", amount.ToString());
                    writer.WriteEndElement(); // Event
                }

                writer.WriteEndElement(); // Timeline
                writer.WriteEndDocument();
            }

            Console.WriteLine($"XML timeline exported to '{Path.GetFullPath(xmlPath)}'");
        }
    }
}