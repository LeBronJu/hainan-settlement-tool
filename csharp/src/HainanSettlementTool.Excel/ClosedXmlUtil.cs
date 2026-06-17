using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class ClosedXmlUtil
    {
        public static double CellNumber(IXLCell cell)
        {
            if (cell.DataType == XLDataType.Number)
            {
                return cell.GetDouble();
            }

            return TextUtil.N(cell.GetFormattedString());
        }

        public static IXLWorksheet MainSheet(XLWorkbook workbook)
        {
            var named = workbook.Worksheets.FirstOrDefault(ws => ws.Name == LedgerLayout.MainSheetName);
            if (named != null)
            {
                return named;
            }

            var sheet1 = workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Sheet1");
            if (sheet1 != null)
            {
                return sheet1;
            }

            var matched = workbook.Worksheets.FirstOrDefault(ws => TextUtil.S(ws.Cell("A1").GetFormattedString()).Contains("售电结算台账"));
            if (matched != null)
            {
                return matched;
            }

            throw new InvalidOperationException("找不到台账主表：" + LedgerLayout.MainSheetName);
        }

        public static string ColumnLetter(int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = string.Empty;
            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        public static string DefaultLedgerOutputName(string baseLedger, int month)
        {
            var stem = Path.GetFileNameWithoutExtension(baseLedger);
            var extension = Path.GetExtension(baseLedger);
            return stem + "】补" + month + "月电量" + extension;
        }
    }
}
