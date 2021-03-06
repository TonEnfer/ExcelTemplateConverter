﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Diagnostics;
using System.IO;
using DocumentFormat.OpenXml.Validation;

namespace ExcelTemplateConverterLib
{
    public sealed class Report
    {
        private string path = null;

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

                    Sheet sheet = new Sheet()
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Report"
                    };
                    sheets.Append(sheet);

                    Columns columns = new Columns();
                    for (uint i = 0; i < 19; i++)
                        columns.AppendChild(new Column
                        {
                            Min = i + 1,
                            Max = i + 1,
                            Width = 15,
                            BestFit = true,
                            //CustomWidth = true

                        });

                    worksheet.AppendChild(columns);

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
                        ExcelUtils.ConstructCell("Счет", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Наименование", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Инвентарный номер", CellValues.String, 2),
                        ExcelUtils.ConstructCell("КФО", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Сальдо на начало периода", CellValues.String, 2),
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 },
                        ExcelUtils.ConstructCell("Обороты за период", CellValues.String, 2),
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 },
                        ExcelUtils.ConstructCell("Сальдо на конец периода", CellValues.String, 2),
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 },
                        ExcelUtils.ConstructCell("Расположение", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Комментарий", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Дата обновления", CellValues.String, 2)
                        );
                    sheetData.Append(row);
                    row = new Row();
                    row.Append(
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 },
                        ExcelUtils.ConstructCell("Дебет", CellValues.String, 2),
                        new Cell() { StyleIndex = 2 },
                        ExcelUtils.ConstructCell("Кредит", CellValues.String, 2),
                        new Cell() { StyleIndex = 2 },
                        ExcelUtils.ConstructCell("Дебет", CellValues.String, 2),
                        new Cell() { StyleIndex = 2 },
                        ExcelUtils.ConstructCell("Кредит", CellValues.String, 2),
                        new Cell() { StyleIndex = 2 },
                        ExcelUtils.ConstructCell("Дебет", CellValues.String, 2),
                        new Cell() { StyleIndex = 2 },
                        ExcelUtils.ConstructCell("Кредит", CellValues.String, 2),
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 }
                        );
                    sheetData.Append(row);

                    row = new Row();
                    row.Append(
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 },
                        ExcelUtils.ConstructCell("Сумма", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Количество", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Сумма", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Количество", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Сумма", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Количество", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Сумма", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Количество", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Сумма", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Количество", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Сумма", CellValues.String, 2),
                        ExcelUtils.ConstructCell("Количество", CellValues.String, 2),
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 },
                        new Cell() { StyleIndex = 2 });
                    sheetData.Append(row);



                    List<MergeCell> mergeCellsList = new List<MergeCell>() {
                        new MergeCell { Reference = "A1:A3" },
                        new MergeCell {Reference = "B1:B3"},
                        new MergeCell {Reference = "C1:C3"},
                        new MergeCell {Reference = "D1:D3"},
                        new MergeCell {Reference = "E1:H1"},
                        new MergeCell {Reference = "I1:L1"},
                        new MergeCell {Reference = "M1:P1"},
                        new MergeCell {Reference = "Q1:Q3"},
                        new MergeCell {Reference = "R1:R3"},
                        new MergeCell {Reference = "S1:S3"},
                        new MergeCell {Reference = "E2:F2"},
                        new MergeCell {Reference = "G2:H2"},
                        new MergeCell {Reference = "I2:J2"},
                        new MergeCell {Reference = "K2:L2"},
                        new MergeCell {Reference = "M2:N2"},
                        new MergeCell {Reference = "O2:P2"}

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

        private Row CreateRowFromDataset(Dataset dataset, uint FormatId = 1)
        {
            Row row = new Row();

            row.Append(
                ExcelUtils.ConstructCell(dataset.Invoice, CellValues.String, FormatId),
                ExcelUtils.ConstructCell(dataset.Name, CellValues.String, FormatId),
                ExcelUtils.ConstructCell(dataset.InventoryNumber,
                dataset.InventoryNumber.All(Char.IsDigit) ? CellValues.Number : CellValues.String, FormatId),
                ExcelUtils.ConstructCell(dataset.KFO, CellValues.Number, FormatId),
                ExcelUtils.ConstructCell(Convert.ToString(dataset.startPeriodBalance.debit.sum).Replace(',', '.'), FormatId),
                ExcelUtils.ConstructCell(Convert.ToString(dataset.startPeriodBalance.debit.amount).Replace(',', '.'), FormatId),
                ExcelUtils.ConstructCell(Convert.ToString(dataset.startPeriodBalance.credit.sum).Replace(',', '.'), FormatId),
                ExcelUtils.ConstructCell(Convert.ToString(dataset.startPeriodBalance.credit.amount).Replace(',', '.'), FormatId),
                ExcelUtils.ConstructCell(Convert.ToString(dataset.turnover.debit.sum).Replace(',', '.'), FormatId),
                ExcelUtils.ConstructCell(Convert.ToString(dataset.turnover.debit.amount).Replace(',', '.'), FormatId),
                ExcelUtils.ConstructCell(Convert.ToString(dataset.turnover.credit.sum).Replace(',', '.'), FormatId),
                ExcelUtils.ConstructCell(Convert.ToString(dataset.turnover.credit.amount).Replace(',', '.'), FormatId),
                ExcelUtils.ConstructCell(Convert.ToString(dataset.endPeriodBalance.debit.sum).Replace(',', '.'), FormatId),
                ExcelUtils.ConstructCell(Convert.ToString(dataset.endPeriodBalance.debit.amount).Replace(',', '.'), FormatId),
                ExcelUtils.ConstructCell(Convert.ToString(dataset.endPeriodBalance.credit.sum).Replace(',', '.'), FormatId),
                ExcelUtils.ConstructCell(Convert.ToString(dataset.endPeriodBalance.credit.amount).Replace(',', '.'), FormatId),
                new Cell() { StyleIndex = FormatId },
                new Cell() { StyleIndex = FormatId },
                new Cell() { StyleIndex = FormatId }
                );
            return row;
        }

        public void WriteDataToExcelDoc(List<Dataset> datasets)
        {
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
                {
                    WorksheetPart worksheetPart = document.WorkbookPart.WorksheetParts.FirstOrDefault();
                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    Stylesheet stylesheet = document.WorkbookPart.WorkbookStylesPart.Stylesheet;
                    CellStyles cellstyles = stylesheet.CellStyles;
                    //var FormatId = ((CellStyle)cellstyles.ChildElements.ToList().Where(a=>((CellStyle)a).Name == "good").FirstOrDefault()).FormatId;
                    //datasets.Where(a => a.Name.Contains("Расходомер"))
                    foreach (var ds in datasets)
                        sheetData.AppendChild(CreateRowFromDataset(ds,3));

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

        public void ExcelValidate()
        {
            try

            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
                {
                    OpenXmlValidator validator = new OpenXmlValidator();

                    int count = 0;

                    foreach (ValidationErrorInfo error in validator.Validate(document))
                    {

                        count++;
                        Console.WriteLine("Error " + count);
                        Console.WriteLine("Description: " + error.Description);
                        Console.WriteLine("Path: " + error.Path.XPath);
                        Console.WriteLine("Part: " + error.Part.Uri);
                        Console.WriteLine("-------------------------------------------");
                    }
                    Console.ReadKey();
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
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
