using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;

namespace Count
{
    public class ExcelBuilder
    {
        public byte[] BuildExcel(
            GatheredInformation gatheredInformation,
            ExtractedValues extractedValues,
            decimal ignitisTotal,
            decimal sumWithoutTelia,
            decimal monthly)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            XSSFCellStyle headerCellStyle = CreateHeaderCellStyle(workbook);

            int rowIndex = 0;
            CreateRowDecimal(sheet, rowIndex++, null, "IGNITIS", ignitisTotal, gatheredInformation.IgnitisTo);
            CreateRowDecimal(sheet, rowIndex++, null, "VV", extractedValues.Vv);
            CreateRowDecimal(sheet, rowIndex++, null, "MANOBUSTAS", extractedValues.ManoBustas);
            CreateRowDecimal(sheet, rowIndex++, null, "VST", extractedValues.Vst);
            CreateRowWithFormula(sheet, rowIndex++, headerCellStyle, "SUBTOTAL_1", "SUM(B1, B2, B3, B4)"); //5
            CreateRowDecimal(sheet, rowIndex++, null, "TELIA", Configuration.ValueTelia); //6
            CreateRowDecimal(sheet, rowIndex++, null, "GARBAGE", Configuration.Garbage); //7
            string grandTotalFormula = "SUM(B5, B6, B7)";
            CreateRowWithFormula(sheet, rowIndex++, headerCellStyle, "SUBTOTAL_2", grandTotalFormula); //8

            IRow blankRow9 = sheet.CreateRow(rowIndex++); //9

            CreateRowDecimal(sheet, rowIndex++, null, "UŽ PRAĖJUSIUS MĖN.", gatheredInformation.DebtFromPreviousMonths); //10
            decimal debtFromThisMonth = (sumWithoutTelia + Configuration.ValueTelia + Configuration.Garbage) % Configuration.RoundToAmount;
            CreateRowDecimal(sheet, rowIndex++, null, "UŽ ŠĮ MĖN.", debtFromThisMonth); //11
            string debtTotalFormula = "SUM(B10, B11)";
            CreateRowWithFormula(sheet, rowIndex++, null, "VISO UŽ MĖN.", debtTotalFormula); //12
            IRow row13 = CreateRowWithFormula(sheet, rowIndex++, null, "VISO UŽ MĖN. KOREG.", debtTotalFormula);

            IRow blankRow14 = sheet.CreateRow(rowIndex++);

            string grandTotalFormulaWithoutDebtFromThisMonth = grandTotalFormula + " - B11";
            IRow row15 = CreateRowWithFormula(sheet, rowIndex++, headerCellStyle, "SUBTOTAL_3", grandTotalFormulaWithoutDebtFromThisMonth);

            if (gatheredInformation.DebtFromPreviousMonths + debtFromThisMonth >= Configuration.RoundToAmount)
            {
                debtTotalFormula = $"{debtTotalFormula} - {Configuration.RoundToAmount}";
                //overwrite current value of second cell
                CreateCellWithFormula(row13, 1, debtTotalFormula);

                grandTotalFormulaWithoutDebtFromThisMonth = $"{grandTotalFormulaWithoutDebtFromThisMonth} + {Configuration.RoundToAmount}";
                //overwrite current value of second cell
                CreateCellWithFormula(row15, 1, grandTotalFormulaWithoutDebtFromThisMonth, headerCellStyle);
            }

            CreateRowDecimal(sheet, rowIndex++, null, "MONTHLY", monthly); //16
            CreateRowWithFormula(sheet, rowIndex, headerCellStyle, "GRANDTOTAL", "SUM(B15, B16)"); //17
            
            sheet.SetColumnWidth(0, 6400);
            sheet.SetColumnWidth(1, 6400);

            byte[] bytes;
            using (ByteArrayOutputStream bos = new ByteArrayOutputStream())
            {
                workbook.Write(bos);
                bytes = bos.ToByteArray();
            }

            return bytes;
        }

        private void CreateRowDecimal(ISheet sheet, int rowIndex, XSSFCellStyle style, string label, decimal value, decimal value2 = decimal.MinValue)
        {
            IRow row = sheet.CreateRow(rowIndex);
            CreateCell(row, 0, label, style);
            CreateCell(row, 1, double.Parse(value.ToString()), style);

            if (value2 != decimal.MinValue)
            {
                CreateCell(row, 2, double.Parse(value2.ToString()), style);
            }
        }

        private IRow CreateRowWithFormula(ISheet sheet, int rowIndex, XSSFCellStyle style, string label, string formula)
        {
            IRow row = sheet.CreateRow(rowIndex);
            CreateCell(row, 0, label, style);
            CreateCellWithFormula(row, 1, formula, style);

            return row;
        }

        private void CreateCell(IRow currentRow, int cellIndex, double value, XSSFCellStyle style = null)
        {
            ICell cell = currentRow.CreateCell(cellIndex);
            cell.SetCellValue(value);
            cell.CellStyle = style;
        }

        private void CreateCell(IRow currentRow, int cellIndex, string value, XSSFCellStyle style = null)
        {
            ICell cell = currentRow.CreateCell(cellIndex);
            cell.SetCellValue(value);
            cell.CellStyle = style;
        }

        private void CreateCellWithFormula(IRow currentRow, int cellIndex, string formula, XSSFCellStyle style = null)
        {
            ICell cell = currentRow.CreateCell(cellIndex);
            cell.SetCellFormula(formula);
            cell.CellStyle = style;
        }

        private XSSFCellStyle CreateHeaderCellStyle(XSSFWorkbook workbook)
        {
            XSSFFont font = (XSSFFont)workbook.CreateFont();
            font.FontHeightInPoints = 11;
            font.FontName = "Tahoma";
            font.IsBold = true;

            XSSFCellStyle headerCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            headerCellStyle.SetFont(font);
            return headerCellStyle;
        }
    }
}