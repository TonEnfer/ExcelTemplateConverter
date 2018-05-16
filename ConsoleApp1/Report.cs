using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Diagnostics;
using System.IO;

namespace ConsoleApp1
{
    public sealed class Report
    {
        string path = null;
        public enum ReportFormat
        {
            ReportWhithInventoryNumber,
            ReportWithoutInventoryNumber
        }

        public Report(string path)
        {
            this.path = path;
        }

        public void CreateExcelDoc()
        {
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    Worksheet worksheet = worksheetPart.Worksheet = new Worksheet();
                    //worksheetPart.Worksheet.Save();

                    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                    Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1, Name = "Report" };
                    sheets.Append(sheet);
                    workbookPart.Workbook.Save();

                    WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                    stylePart.Stylesheet = ExcelUtils.GenerateStylesheet();
                    stylePart.Stylesheet.Save();
                    
                    SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());
                    Row row = new Row();
                    MergeCells mergeCells;
                    if (worksheet.Elements<MergeCells>().Count() > 0)
                        mergeCells = worksheet.Elements<MergeCells>().First();
                    else
                    {
                        mergeCells = new MergeCells();
                        if (worksheet.Elements<CustomChartsheetView>().Count() > 0)
                            worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<CustomChartsheetView>().First());
                        else
                            worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.GetFirstChild<SheetData>());
                    }
                    row.Append(
                        ExcelUtils.ConstructCell("Счёт", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Наименование", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Инвентарный номер", CellValues.String, 2),
                        ExcelUtils.ConstructCell("КФО", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Сальдо на начало периода", CellValues.String, 2),
                        ExcelUtils.ConstructCell("", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Обороты за период", CellValues.String, 2),
                        ExcelUtils.ConstructCell("", CellValues.String, 1),
                        ExcelUtils.ConstructCell("Сальдо на конец периода", CellValues.String, 2),
                        ExcelUtils.ConstructCell("", CellValues.String, 0),
                        ExcelUtils.ConstructCell("Расположение", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Комментарий", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Дата обновления", CellValues.String, 2)
                        );
                    sheetData.Append(row);
                    row = new Row();
                    row.Append(
                        ExcelUtils.ConstructCell("", CellValues.Number),
                        ExcelUtils.ConstructCell("", CellValues.Number),
                        ExcelUtils.ConstructCell("", CellValues.Number),
                        ExcelUtils.ConstructCell("", CellValues.Number),
                        ExcelUtils.ConstructCell("Дебет", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Кредит", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Дебет", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Кредит", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Дебет", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Кредит", CellValues.String, 2));
                    sheetData.Append(row);
                    //Console.WriteLine(row.GetFirstChild<Cell>().CellReference);
                    List<MergeCell> mergeCellsList = new List<MergeCell>() {
                        new MergeCell { Reference = "A1:A2" },
                        new MergeCell {Reference = "B1:B2"},
                        new MergeCell {Reference = "C1:C2"},
                        new MergeCell {Reference = "D1:D2"},
                        new MergeCell {Reference = "E1:F1"},
                        new MergeCell {Reference = "G1:H1"},
                        new MergeCell {Reference = "I1:J1"},
                        new MergeCell {Reference = "K1:K2"},
                        new MergeCell {Reference = "L1:L2"},
                        new MergeCell {Reference = "M1:M2"}

                    };
                    mergeCells.Append(mergeCellsList);

                    worksheetPart.Worksheet.Save();

                }
            }
            catch (Exception e)
            {
                Debug.Print(e.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

        public void WriteExcelDoc()
        {
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
                {
                    WorksheetPart worksheetPart = document.WorkbookPart.WorksheetParts.FirstOrDefault();
                    Worksheet worksheet = worksheetPart.Worksheet;
                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    Row row = new Row();
                    MergeCells mergeCells;
                    if (worksheet.Elements<MergeCells>().Count() > 0)
                        mergeCells = worksheet.Elements<MergeCells>().First();
                    else
                    {
                        mergeCells = new MergeCells();
                        if (worksheet.Elements<CustomChartsheetView>().Count() > 0)
                            worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomChartsheetView>().First());
                        else
                            worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                    }
                    row.Append(
                        ExcelUtils.ConstructCell("Счёт", CellValues.String),
                        ExcelUtils.ConstructCell("Наименование", CellValues.String),
                        ExcelUtils.ConstructCell("Инвентарный номер", CellValues.String),
                        ExcelUtils.ConstructCell("КФО", CellValues.String),
                        ExcelUtils.ConstructCell("Сальдо на начало периода", CellValues.String),
                        ExcelUtils.ConstructCell("", CellValues.String),
                        ExcelUtils.ConstructCell("Обороты за период", CellValues.String),
                        ExcelUtils.ConstructCell("", CellValues.String),
                        ExcelUtils.ConstructCell("Сальдо на конец периода", CellValues.String),
                        ExcelUtils.ConstructCell("", CellValues.String),
                        ExcelUtils.ConstructCell("Расположение", CellValues.String),
                        ExcelUtils.ConstructCell("Комментарий", CellValues.String),
                        ExcelUtils.ConstructCell("Дата обновления", CellValues.String)
                        );
                    sheetData.Append(row);
                    row = new Row();
                    row.Append(
                        ExcelUtils.ConstructCell("", CellValues.Number),
                        ExcelUtils.ConstructCell("", CellValues.Number),
                        ExcelUtils.ConstructCell("", CellValues.Number),
                        ExcelUtils.ConstructCell("", CellValues.Number),
                        ExcelUtils.ConstructCell("Дебет", CellValues.String),
                        ExcelUtils.ConstructCell("Кредит", CellValues.String),
                        ExcelUtils.ConstructCell("Дебет", CellValues.String),
                        ExcelUtils.ConstructCell("Кредит", CellValues.String),
                        ExcelUtils.ConstructCell("Дебет", CellValues.String),
                        ExcelUtils.ConstructCell("Кредит", CellValues.String));
                    sheetData.Append(row);
                    //Console.WriteLine(row.GetFirstChild<Cell>().CellReference);
                    List<MergeCell> mergeCellsList = new List<MergeCell>() {
                        new MergeCell { Reference = "A1:A2" },
                        new MergeCell {Reference = "B1:B2"},
                        new MergeCell {Reference = "C1:C2"},
                        new MergeCell {Reference = "D1:D2"},
                        new MergeCell {Reference = "E1:F1"},
                        new MergeCell {Reference = "G1:H1"},
                        new MergeCell {Reference = "I1:J1"},
                        new MergeCell {Reference = "K1:K2"},
                        new MergeCell {Reference = "L1:L2"},
                        new MergeCell {Reference = "M1:M2"}

                    };
                    mergeCells.Append(mergeCellsList);

                    worksheetPart.Worksheet.Save();

                }
            }
            catch (Exception e)
            {
                Debug.Print(e.Message);
            }
            finally
            {
                GC.Collect();
            }
        }



        public List<Dataset> ReadExcelDoc()
        {
            List<Dataset> datasets = null;
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
                {
                    datasets = ExcelParser.GetParser(document).Parse(document);
                    return datasets;
                }
            }
            catch (Exception e)
            {
                Debug.Print(e.Message);
                Debug.Print(e.StackTrace);
                return null;
            }
            finally
            {

                GC.Collect();
            }
        }


    }
}
