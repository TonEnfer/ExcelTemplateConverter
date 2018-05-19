using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;

namespace ExcelTemplateConverterLib
{
    public static class ExcelUtils
    {
        public static int FindStringId(SharedStringTablePart sharedStrings, string value)
        {
            int markerId = -100;
            //sharedStrings.SharedStringTable.ChildElements.ToList().IndexOf(
            //    sharedStrings.SharedStringTable.ChildElements.Where(k => k.InnerText == value).FirstOrDefault());
            if (sharedStrings != null)
                for (int i = 0; i < sharedStrings.SharedStringTable.ChildElements.Count; i++)
                {
                    if (sharedStrings.SharedStringTable.ChildElements[i].InnerText == value)
                    {
                        markerId = i;
                        break;
                    }
                }

            return markerId;
        }
        public static string FindStringValue(SharedStringTablePart sharedStrings, int id)
        {
            return sharedStrings.SharedStringTable.ChildElements[id].InnerText;
        }

        public static Row GetRow(SheetData wsData, UInt32 rowIndex)
        {
            var row = wsData.Elements<Row>().
            Where(r => r.RowIndex.Value == rowIndex).FirstOrDefault();
            if (row == null)
            {
                row = new Row() { RowIndex = rowIndex };
                wsData.Append(row);
            }
            return row;
        }
        public static Cell GetCell(Row r, string collumnName)
        {

            return r.Elements<Cell>()
                .Where(c => (Regex.IsMatch(c.CellReference.Value, collumnName + "[0-9]*")))
                .FirstOrDefault();
        }
        public static string GetCellText(Row r, string col)
        {
            Cell c = GetCell(r, col);
            if (c == null)
                return "0";
            if (c.CellValue == null)
                return "0";

            return c.CellValue.Text;
        }


        public static Worksheet GetWorksheet(SpreadsheetDocument document, string worksheetName)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook
                .Descendants<Sheet>().Where(s => s.Name == worksheetName);
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart
                .GetPartById(sheets.First().Id);
            return worksheetPart.Worksheet;
        }

        public static Cell ConstructCell(string value, CellValues dataType, uint styleIndex = 0)
        {
            return new Cell()
            {
                CellValue = value == "0" ? new CellValue() : new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
                StyleIndex = styleIndex
            };
        }

        public static Cell ConstructCell(string value, uint styleIndex = 0)
        {
            return new Cell()
            {
                CellValue = value == "0" ? new CellValue() : new CellValue(value),
                StyleIndex = styleIndex
            };
        }

        public static Stylesheet GenerateStylesheet()
        {
            Stylesheet styleSheet = null;



            Fonts fonts = new Fonts(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 },
                    new FontName() { Val = "Times New Roman" }

                ),
                new Font( // Index 1 - header
                    new FontSize() { Val = 10 },
                    new Bold(),
                    new Color() { Rgb = "FFFFFFFF" },
                    new FontName() { Val = "Times New Roman" }



                ));

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "FF4472C4" } })
                    { PatternType = PatternValues.Solid }), // Index 2 - header
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "FFFF0000" } })
                    { PatternType = PatternValues.Solid }), //Index 3 - bad
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "FF00FF00" } })
                    { PatternType = PatternValues.Solid }) //index 4 - good
                );

            Borders borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );
            CellStyleFormats cellStyleFormats = new CellStyleFormats(
                new CellFormat(),
                new CellFormat()
                {
                    FontId = 1,
                    FillId = 3,
                    BorderId = 1,
                    ApplyFill = true,
                    FormatId = 0,
                }, // 1
                new CellFormat()
                {
                    FontId = 1,
                    FillId = 4,
                    BorderId = 1,
                    ApplyFill = true,
                    FormatId = 1
                } // 2
                );

            CellFormats cellFormats = new CellFormats(
                    new CellFormat() { FormatId = 0 }, //0 default
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, //1 body
                    new CellFormat(
                        new Alignment()
                        {
                            Horizontal = HorizontalAlignmentValues.CenterContinuous,
                            Vertical = VerticalAlignmentValues.Center
                        })
                    { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true }, //2 
                    new CellFormat() {FormatId = 1 } // 3



                );
            CellStyles cellStyles = new CellStyles(
                new CellStyle() { FormatId = 0 },
                new CellStyle() { Name = "good", FormatId = 2, BuiltinId = 26 },
                new CellStyle() { Name = "bad", FormatId = 1, BuiltinId = 27 }
                );

            styleSheet = new Stylesheet(fonts, fills, borders, cellStyleFormats, cellFormats, cellStyles);

            return styleSheet;
        }
    }
}
