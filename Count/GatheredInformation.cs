using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Count
{
    public class GatheredInformation
    {
        public GatheredInformation(string customerNumber)
        {
            DateTime previousDate = DateTime.Now.AddMonths(-2);
            string path = $"{Configuration.DownloadsPath}\\{customerNumber}_{previousDate.Year}_{previousDate.Month}.xlsx";

            XSSFWorkbook xssfWorkbook;
            using (FileStream previousMonthFile = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                xssfWorkbook = new XSSFWorkbook(previousMonthFile);
            }

            ISheet sheet = xssfWorkbook.GetSheet("Sheet1");
            IgnitisFrom = int.Parse(sheet.GetRow(0).GetCell(2).NumericCellValue.ToString());

            IRow debtFromPreviousMonthsRow = sheet.GetRow(12);
            ICell debtFromPreviousMonthsCell = debtFromPreviousMonthsRow.GetCell(1);
            var formula = new XSSFFormulaEvaluator(xssfWorkbook);
            formula.EvaluateInCell(debtFromPreviousMonthsCell);
            DebtFromPreviousMonths = decimal.Parse(debtFromPreviousMonthsCell.NumericCellValue.ToString());

            SumToValidate = decimal.Parse(File.ReadAllText($"{Configuration.DownloadsPath}\\New Text Document.txt"), Configuration.Nfi);

            Console.WriteLine("Ignitis iki:");
            IgnitisTo = int.Parse(Console.ReadLine());
        }

        public decimal DebtFromPreviousMonths { get; set; }
        public int IgnitisFrom { get; set; }
        public int IgnitisTo { get; set; }
        public decimal SumToValidate { get; set; }
    }
}