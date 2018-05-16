using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Diagnostics;

namespace ConsoleApp1
{
    class Program
    {
        const string path = "E:\\";
        const string fileName2 = "по счету 101.00 за 15 марта 2018 г_.xlsx";
        const string fileName3 = "по счету 29 за 15 марта 2018 г_.xlsx";
        const string fileName = "test.xlsx";


        static void Main(string[] args)
        {
            Report newReport = new Report(Path.Combine(path, fileName)),
                oldReport = new Report(Path.Combine(path, fileName2)),
                oldReport2 = new Report(Path.Combine(path, fileName3));

            newReport.CreateExcelDoc();
            //List<Dataset> datasets1 = oldReport.ReadExcelDoc();


            List<Dataset> datasets2 = oldReport2.ReadExcelDoc();
            newReport.WriteDataToExcelDoc(datasets2);
            newReport.ExcelValidate();

            Console.ReadKey();
        }
    }
}
