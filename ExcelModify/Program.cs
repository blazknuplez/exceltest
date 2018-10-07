using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelModify
{
    class Program
    {
        static void Main(string[] args)
        {
            string documentPath = "Data//Sample.xlsx";
            OpenExcelDocument(documentPath);

            Console.Write("Done");
        }

        public static void OpenExcelDocument(string docName)
        {
            List<string> values = new List<string>
            {
                "Naslov 1", "Naslov 2", "Naslov 3", "Naslov 4", "Naslov 5"
            };

            int dataSize = 20;

            byte[] byteArray = File.ReadAllBytes(docName);
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, (int)byteArray.Length);

                using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true))
                {
                    WorkbookPart wbPart = document.WorkbookPart;
                    WorksheetPart worksheetPart = wbPart.WorksheetParts.First();

                    CreateColumnsIfNecessary(worksheetPart, values.Count);

                    SheetData sheetdata = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    List<UInt32> footerStyles = GetFooterStyles(sheetdata);

                    CreateHeaderData(sheetdata, values);

                    for (int i = 1; i < dataSize; i++)
                    {
                        CreateBodyData(sheetdata, i, values);
                    }

                    CreateFooterData(sheetdata, values, footerStyles);
                }

                File.WriteAllBytes("Data//Sample_saved.xlsx", stream.ToArray());
            }
        }

        public static void CreateColumnsIfNecessary(WorksheetPart worksheetPart, int count)
        {
            Columns columns = worksheetPart.Worksheet.GetFirstChild<Columns>();

            int difference = count - columns.Elements<Column>().Count();
            Column column = columns.Elements<Column>().LastOrDefault();
            for (int i = 0; i < difference; i++)
            {

                columns.AppendChild(new Column()
                {
                    Min = column.Min + (UInt32)(i + 1),
                    Max = column.Max + (UInt32)(i + 1),
                    Width = column.Width,
                    Style = column.Style,
                    CustomWidth = true
                });
            }
        }

        public static void FillRow(Row row, List<string> values, List<UInt32> styles)
        {
            for (int i = 0; i < values.Count; i++)
            {
                string value = values[i];
                Cell cell = row.Elements<Cell>().ElementAtOrDefault(i);

                bool exist = cell != null;

                if (!exist)
                {
                    cell = new Cell();
                }

                cell.CellValue = new CellValue(value);
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                cell.StyleIndex = i != values.Count - 1
                    ? styles.FirstOrDefault()
                    : styles.LastOrDefault();

                if (!exist)
                {
                    row.AppendChild(cell);
                }
            }
        }

        public static void CreateHeaderData(SheetData sheetdata, List<string> values)
        {
            List<UInt32> headerStyles = GetHeaderStyles(sheetdata);
            var header = sheetdata.Elements<Row>().FirstOrDefault();

            FillRow(header, values, headerStyles);
        }

        public static void CreateBodyData(SheetData sheetdata, int rowIndex, List<string> values)
        {
            List<UInt32> bodyStyles = GetBodyStyles(sheetdata);
            Row row = sheetdata.Elements<Row>().ElementAtOrDefault(rowIndex);

            bool exist = row != null;

            if (!exist)
            {
                row = new Row();
            }

            FillRow(row, values, bodyStyles);

            if (!exist)
            {
                sheetdata.AppendChild(row);
            }
        }

        public static void CreateFooterData(SheetData sheetdata, List<string> values, List<UInt32> footerStyles)
        {
            var footer = sheetdata.Elements<Row>().LastOrDefault();

            FillRow(footer, values, footerStyles);
        }

        public static List<UInt32> GetHeaderStyles(SheetData sheetdata)
        {
            Row header = sheetdata.Elements<Row>().FirstOrDefault();

            return new List<UInt32>
            {
                header.Elements<Cell>().FirstOrDefault()?.StyleIndex,
                header.Elements<Cell>().LastOrDefault()?.StyleIndex
            };
        }

        public static List<UInt32> GetBodyStyles(SheetData sheetdata)
        {
            Row footer = sheetdata.Elements<Row>().Skip(1).FirstOrDefault();

            return new List<UInt32>
            {
                footer.Elements<Cell>().FirstOrDefault()?.StyleIndex,
                footer.Elements<Cell>().LastOrDefault()?.StyleIndex
            };
        }

        public static List<UInt32> GetFooterStyles(SheetData sheetdata)
        {
            Row footer = sheetdata.Elements<Row>().LastOrDefault();

            return new List<UInt32>
            {
                footer.Elements<Cell>().FirstOrDefault()?.StyleIndex,
                footer.Elements<Cell>().LastOrDefault()?.StyleIndex
            };
        }
    }
}
