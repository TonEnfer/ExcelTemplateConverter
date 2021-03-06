﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Diagnostics;

namespace ExcelTemplateConverterLib
{
    public abstract class ExcelParser
    {
        public abstract List<Dataset> Parse(SpreadsheetDocument document);
        protected static uint FindRowIndexByMarker(SpreadsheetDocument document, string marker)
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
            Sheets sheets = workbookPart.Workbook.Sheets;
            Sheet sheet = sheets.GetFirstChild<Sheet>();

            SheetData sheetData = worksheetPart.Worksheet.Descendants<SheetData>().FirstOrDefault();
            SharedStringTablePart sharedStringPart = workbookPart.SharedStringTablePart;
            uint firstRow = 1;
            int startMarkerId = ExcelUtils.FindStringId(sharedStringPart, marker);
            foreach (Row rr in sheetData.Elements<Row>())
            {
                Cell cc = rr.Elements<Cell>().FirstOrDefault();
                if (cc != null)
                    if ((cc.DataType == "s") && (Convert.ToInt32(cc.CellValue.Text) == startMarkerId))
                        break;

                firstRow++;
            }
            return firstRow;
        }
        public static ExcelParser GetParser(SpreadsheetDocument document)
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
            Sheets sheets = workbookPart.Workbook.Sheets;
            Sheet sheet = sheets.GetFirstChild<Sheet>();

            SheetData sheetData = worksheetPart.Worksheet.Descendants<SheetData>().FirstOrDefault();
            SharedStringTablePart sharedStringPart = workbookPart.SharedStringTablePart;

            string startMarker = "Счет";
            int startMarkerID = ExcelUtils.FindStringId(sharedStringPart, startMarker);
            uint firstRow = 1;
            if (startMarkerID != -100)
                firstRow = FindRowIndexByMarker(document, startMarker);
            string formatMarker = "КФО";
            int formatMarkerID = ExcelUtils.FindStringId(sharedStringPart, formatMarker);
            if (startMarkerID != -100 || formatMarkerID != -100)
            {
                Row r = ExcelUtils.GetRow(sheetData, 1);
                if (Convert.ToInt32(ExcelUtils.GetCell(r, "A").CellValue.Text) == startMarkerID ||
                    Convert.ToInt32(ExcelUtils.GetCell(r, "D").CellValue.Text) == formatMarkerID)
                    return new ParserNewFormat();
            }
            else
            {
                Row r = ExcelUtils.GetRow(sheetData, 1);
                if (ExcelUtils.GetCell(r, "A").DataType == CellValues.String &&
                    ExcelUtils.GetCell(r, "A").CellValue.Text == startMarker ||
                    ExcelUtils.GetCell(r, "D").DataType == CellValues.String &&
                    ExcelUtils.GetCell(r, "D").CellValue.Text == formatMarker)
                    return new ParserNewFormat();
            }
            if (formatMarkerID != -100)
                if (Convert.ToInt32(ExcelUtils.GetRow(sheetData, firstRow + 3).GetFirstChild<Cell>().CellValue.Text) == formatMarkerID)
                    return new ParserWhithInventoryNumber();
            return new ParserWhithoutInventoryNumber();

        }
    }
    class ParserNewFormat : ExcelParser
    {
        public override List<Dataset> Parse(SpreadsheetDocument document)
        {
            List<Dataset> datasets = null;
            try
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                Sheets sheets = workbookPart.Workbook.Sheets;
                Sheet sheet = sheets.GetFirstChild<Sheet>();

                SheetData sheetData = worksheetPart.Worksheet.Descendants<SheetData>().FirstOrDefault();
                SharedStringTablePart sharedStringPart = workbookPart.SharedStringTablePart;

                string startMarker = "Счет";
                uint firstRow = 4;
                int lastRow = sheetData.Elements<Row>().Count();

                datasets = new List<Dataset>();

                for (uint i = firstRow; i < lastRow+1; i += 1)
                {
                    Console.WriteLine("Строка {0}", i);
                    Dataset ds = new Dataset();
                    Row r = ExcelUtils.GetRow(sheetData, i);
                    var q = r.Elements<Cell>().Where(c => c.CellValue != null).ToList();
                    ds.Invoice = ExcelUtils.FindStringValue(sharedStringPart, Convert.ToInt32(ExcelUtils.GetCellText(r, "A")));
                    ds.Name = ExcelUtils.FindStringValue(sharedStringPart, Convert.ToInt32(ExcelUtils.GetCellText(r, "B")));
                    ds.InventoryNumber = ExcelUtils.GetCell(r, "C").DataType == "s" ?
                       ExcelUtils.FindStringValue(sharedStringPart, Convert.ToInt32(ExcelUtils.GetCellText(r, "C"))).Trim(' ') :
                       ExcelUtils.GetCellText(r, "C").Trim(' ');
                    ds.KFO = ExcelUtils.GetCellText(r, "D").Trim(' ');
                    ds.startPeriodBalance.debit.sum = Convert.ToDouble(ExcelUtils.GetCellText(r, "E").Replace('.', ','));
                    ds.startPeriodBalance.debit.amount = Convert.ToInt32(ExcelUtils.GetCellText(r, "F").Replace('.', ','));
                    ds.startPeriodBalance.credit.sum = Convert.ToDouble(ExcelUtils.GetCellText(r, "G").Replace('.', ','));
                    ds.startPeriodBalance.credit.amount = Convert.ToInt32(ExcelUtils.GetCellText(r, "H").Replace('.', ','));
                    ds.turnover.debit.sum = Convert.ToDouble(ExcelUtils.GetCellText(r, "I").Replace('.', ','));
                    ds.turnover.debit.amount = Convert.ToInt32(ExcelUtils.GetCellText(r, "J").Replace('.', ','));
                    ds.turnover.credit.sum = Convert.ToDouble(ExcelUtils.GetCellText(r, "K").Replace('.', ','));
                    ds.turnover.credit.amount = Convert.ToInt32(ExcelUtils.GetCellText(r, "L").Replace('.', ','));
                    ds.endPeriodBalance.debit.sum = Convert.ToDouble(ExcelUtils.GetCellText(r, "M").Replace('.', ','));
                    ds.endPeriodBalance.debit.amount = Convert.ToInt32(ExcelUtils.GetCellText(r, "N").Replace('.', ','));
                    ds.endPeriodBalance.credit.sum = Convert.ToInt32(ExcelUtils.GetCellText(r, "O").Replace('.', ','));
                    ds.endPeriodBalance.credit.amount = Convert.ToInt32(ExcelUtils.GetCellText(r, "P").Replace('.', ','));

                    datasets.Add(ds);
                }
                return datasets;
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

    class ParserWhithInventoryNumber : ExcelParser
    {
        public override List<Dataset> Parse(SpreadsheetDocument document)
        {
            List<Dataset> datasets = null;
            try
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                Sheets sheets = workbookPart.Workbook.Sheets;
                Sheet sheet = sheets.GetFirstChild<Sheet>();

                SheetData sheetData = worksheetPart.Worksheet.Descendants<SheetData>().FirstOrDefault();
                SharedStringTablePart sharedStringPart = workbookPart.SharedStringTablePart;

                string startMarker = "Счет";
                uint firstRow = FindRowIndexByMarker(document, startMarker);

                string endMarker = "Итого";
                uint lastRow = FindRowIndexByMarker(document, endMarker);

                string invoice = ExcelUtils.FindStringValue(sharedStringPart,
                    Convert.ToInt32(ExcelUtils.GetRow(sheetData, firstRow + 4).Elements<Cell>().FirstOrDefault().CellValue.Text));

                datasets = new List<Dataset>();

                for (uint i = firstRow + 6; i < lastRow; i += 2)
                {
                    Console.WriteLine("Строка {0}", i);
                    Dataset ds = new Dataset();
                    Row r = ExcelUtils.GetRow(sheetData, i);
                    var q = r.Elements<Cell>().Where(c => c.CellValue != null).ToList<Cell>();
                    ds.Invoice = invoice;
                    ds.Name = ExcelUtils.FindStringValue(sharedStringPart, Convert.ToInt32(ExcelUtils.GetCellText(r, "A")));
                    ds.startPeriodBalance.debit.sum = Convert.ToDouble(ExcelUtils.GetCellText(r, "K").Replace('.', ','));
                    ds.startPeriodBalance.credit.sum = Convert.ToDouble(ExcelUtils.GetCellText(r, "N").Replace('.', ','));
                    ds.turnover.debit.sum = Convert.ToDouble(ExcelUtils.GetCellText(r, "T").Replace('.', ','));
                    ds.turnover.credit.sum = Convert.ToDouble(ExcelUtils.GetCellText(r, "AA").Replace('.', ','));
                    ds.endPeriodBalance.debit.sum = Convert.ToDouble(ExcelUtils.GetCellText(r, "AF").Replace('.', ','));
                    ds.endPeriodBalance.credit.sum = 0.0;

                    r = ExcelUtils.GetRow(sheetData, i += 1);
                    ds.startPeriodBalance.debit.amount = Convert.ToInt32(ExcelUtils.GetCellText(r, "K"));
                    ds.startPeriodBalance.credit.amount = Convert.ToInt32(ExcelUtils.GetCellText(r, "N"));
                    ds.turnover.debit.amount = Convert.ToInt32(ExcelUtils.GetCellText(r, "T"));
                    ds.turnover.credit.amount = Convert.ToInt32(ExcelUtils.GetCellText(r, "AA"));
                    ds.endPeriodBalance.debit.amount = Convert.ToInt32(ExcelUtils.GetCellText(r, "AF"));
                    ds.endPeriodBalance.credit.amount = 0;

                    r = ExcelUtils.GetRow(sheetData, i += 1);
                    ds.InventoryNumber = ExcelUtils.GetCell(r, "A").DataType == "s" ?
                        ExcelUtils.FindStringValue(sharedStringPart, Convert.ToInt32(ExcelUtils.GetCellText(r, "A"))).Trim(' ') :
                        ExcelUtils.GetCellText(r, "A").Trim(' ');

                    r = ExcelUtils.GetRow(sheetData, i += 2);
                    ds.KFO = ExcelUtils.GetCellText(r, "A").Trim(' ');
                    datasets.Add(ds);
                }
                return datasets;
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
    class ParserWhithoutInventoryNumber : ExcelParser
    {
        public override List<Dataset> Parse(SpreadsheetDocument document)
        {
            List<Dataset> datasets = null;
            try
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                Sheets sheets = workbookPart.Workbook.Sheets;
                Sheet sheet = sheets.GetFirstChild<Sheet>();

                SheetData sheetData = worksheetPart.Worksheet.Descendants<SheetData>().FirstOrDefault();
                SharedStringTablePart sharedStringPart = workbookPart.SharedStringTablePart;

                string startMarker = "Счет";
                uint firstRow = FindRowIndexByMarker(document, startMarker);

                string endMarker = "Итого";
                uint lastRow = FindRowIndexByMarker(document, endMarker);

                string invoice = ExcelUtils.GetRow(sheetData, firstRow + 2).Elements<Cell>().FirstOrDefault().DataType == "s" ?
                    ExcelUtils.FindStringValue(sharedStringPart,
                    Convert.ToInt32(ExcelUtils.GetRow(sheetData, firstRow + 2).Elements<Cell>().FirstOrDefault().CellValue.Text)) :
                    ExcelUtils.GetRow(sheetData, firstRow + 2).Elements<Cell>().FirstOrDefault().CellValue.Text;

                datasets = new List<Dataset>();

                for (uint i = firstRow + 4; i < lastRow; i++)
                {
                    Console.WriteLine("Строка {0}", i);
                    Dataset ds = new Dataset();
                    Row r = ExcelUtils.GetRow(sheetData, i);
                    var q = r.Elements<Cell>().Where(c => c.CellValue != null).ToList();
                    ds.Invoice = invoice;
                    ds.Name = ExcelUtils.FindStringValue(sharedStringPart, Convert.ToInt32(ExcelUtils.GetCellText(r, "A")));
                    ds.startPeriodBalance.debit.sum = Convert.ToDouble(ExcelUtils.GetCellText(r, "K").Replace('.', ','));
                    ds.startPeriodBalance.credit.sum = Convert.ToDouble(ExcelUtils.GetCellText(r, "N").Replace('.', ','));
                    ds.turnover.debit.sum = Convert.ToDouble(ExcelUtils.GetCellText(r, "T").Replace('.', ','));
                    ds.turnover.credit.sum = Convert.ToDouble(ExcelUtils.GetCellText(r, "AA").Replace('.', ','));
                    ds.endPeriodBalance.debit.sum = Convert.ToDouble(ExcelUtils.GetCellText(r, "AF").Replace('.', ','));
                    ds.endPeriodBalance.credit.sum = 0.0;

                    r = ExcelUtils.GetRow(sheetData, i += 1);
                    ds.startPeriodBalance.debit.amount = Convert.ToDouble(ExcelUtils.GetCellText(r, "K").Replace('.', ','));
                    ds.startPeriodBalance.credit.amount = Convert.ToDouble(ExcelUtils.GetCellText(r, "N").Replace('.', ','));
                    ds.turnover.debit.amount = Convert.ToDouble(ExcelUtils.GetCellText(r, "T").Replace('.', ','));
                    ds.turnover.credit.amount = Convert.ToDouble(ExcelUtils.GetCellText(r, "AA").Replace('.', ','));
                    ds.endPeriodBalance.debit.amount = Convert.ToDouble(ExcelUtils.GetCellText(r, "AF").Replace('.', ','));
                    ds.endPeriodBalance.credit.amount = 0;


                    datasets.Add(ds);
                }
                return datasets;
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
