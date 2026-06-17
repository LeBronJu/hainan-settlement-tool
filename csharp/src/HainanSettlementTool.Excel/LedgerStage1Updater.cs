using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Newtonsoft.Json;

namespace HainanSettlementTool.Excel
{
    internal sealed class LedgerStage1Updater
    {
        public Stage1Report Update(Stage1Options options, IStage1ExcelGateway gateway)
        {
            Directory.CreateDirectory(options.OutputDirectory);
            var outputPath = Path.Combine(
                options.OutputDirectory,
                string.IsNullOrWhiteSpace(options.OutputLedgerName)
                    ? ClosedXmlUtil.DefaultLedgerOutputName(options.BaseLedgerPath, options.Month)
                    : options.OutputLedgerName);
            File.Copy(options.BaseLedgerPath, outputPath, true);

            var reportPath = Path.Combine(options.OutputDirectory, options.Month + "月台账更新报告.json");
            var powerRows = gateway.ReadPowerRows(options.PowerPath);
            var rawCodeMap = gateway.ReadCustomerCodes(options.RawDetailPath);

            using (var workbook = new XLWorkbook(outputPath))
            {
                var worksheet = ClosedXmlUtil.MainSheet(workbook);
                var ledgerMap = LedgerRowMap(worksheet);
                var refWorkbook = !string.IsNullOrWhiteSpace(options.ReferenceLedgerPath) && File.Exists(options.ReferenceLedgerPath)
                    ? new XLWorkbook(options.ReferenceLedgerPath)
                    : null;

                try
                {
                    var refSheet = refWorkbook == null ? null : ClosedXmlUtil.MainSheet(refWorkbook);
                    var refMap = refSheet == null ? new Dictionary<string, int>() : LedgerRowMap(refSheet);
                    var lastDataRow = ledgerMap.Values.Max();
                    var nextRow = lastDataRow + 1;
                    var nextSeq = ledgerMap.Values.Select(row => (int)TextUtil.N(worksheet.Cell(row, 1).Value)).DefaultIfEmpty(0).Max() + 1;
                    var targetStart = LedgerLayout.MonthStartColumn(options.Month);
                    var targetMonthAlreadyPresent = TextUtil.S(worksheet.Cell(1, targetStart).GetFormattedString()) == options.Month + "月";

                    if (!targetMonthAlreadyPresent)
                    {
                        CopyMonthBlock(worksheet, options.Month - 1, options.Month);
                    }

                    var report = CreateReport(options, outputPath, reportPath, targetStart, targetMonthAlreadyPresent, powerRows, rawCodeMap);

                    foreach (var item in powerRows)
                    {
                        var foundExisting = ledgerMap.ContainsKey(item.Key);
                        int targetRow;
                        if (foundExisting)
                        {
                            targetRow = ledgerMap[item.Key];
                            report.MatchedCustomers.Add(new RowMatchReport { Name = item.Name, TargetRow = targetRow, Total = item.Total });
                        }
                        else
                        {
                            targetRow = nextRow++;
                            CopyRowTemplate(worksheet, lastDataRow, targetRow, options.Month);
                            worksheet.Cell(targetRow, 1).Value = nextSeq++;
                            worksheet.Cell(targetRow, 3).Value = item.Name;
                            ledgerMap[item.Key] = targetRow;
                            report.NewCustomers.Add(new RowMatchReport { Name = item.Name, TargetRow = targetRow, Total = item.Total });
                        }

                        CopyReferenceIfNeeded(options, worksheet, refSheet, refMap, item, foundExisting, targetRow, report);
                        FillCodeIfMissing(worksheet, rawCodeMap, item, targetRow, report);
                        WriteMonthPower(worksheet, targetStart, targetRow, item);
                    }

                    FillPostWriteChecks(worksheet, ledgerMap, powerRows, report);

                    workbook.SaveAs(outputPath);
                    File.WriteAllText(reportPath, JsonConvert.SerializeObject(report, Formatting.Indented), System.Text.Encoding.UTF8);
                    return report;
                }
                finally
                {
                    refWorkbook?.Dispose();
                }
            }
        }

        private static Stage1Report CreateReport(
            Stage1Options options,
            string outputPath,
            string reportPath,
            int targetStart,
            bool targetMonthAlreadyPresent,
            IList<PowerRow> powerRows,
            IDictionary<string, string> rawCodeMap)
        {
            return new Stage1Report
            {
                Month = options.Month,
                SourceLedger = options.BaseLedgerPath,
                SourcePower = options.PowerPath,
                RawDetailForCodes = options.RawDetailPath,
                ReferenceLedger = options.ReferenceLedgerPath,
                Output = outputPath,
                ReportPath = reportPath,
                TargetBlock = ClosedXmlUtil.ColumnLetter(targetStart) + ":" + ClosedXmlUtil.ColumnLetter(targetStart + LedgerLayout.MonthBlockWidth - 1),
                TargetMonthAlreadyPresent = targetMonthAlreadyPresent,
                PowerRows = powerRows.Count,
                RawDetailCodeRows = rawCodeMap.Count,
                MonthTotal = System.Math.Round(powerRows.Sum(row => row.Total), 4)
            };
        }

        private static void CopyReferenceIfNeeded(
            Stage1Options options,
            IXLWorksheet output,
            IXLWorksheet reference,
            IDictionary<string, int> referenceMap,
            PowerRow item,
            bool foundExisting,
            int targetRow,
            Stage1Report report)
        {
            if (reference == null)
            {
                return;
            }

            int refRow;
            var shouldCopyRef = referenceMap.TryGetValue(item.Key, out refRow) && (!foundExisting || options.CopyReferenceExisting);
            if (shouldCopyRef)
            {
                CopyBaseInfoFromReference(output, reference, targetRow, refRow);
                if (output.Cell(targetRow, 1).IsEmpty())
                {
                    output.Cell(targetRow, 1).Value = targetRow - 3;
                }
                output.Cell(targetRow, 3).Value = item.Name;
                report.CopiedFromReference.Add(new RowMatchReport { Name = item.Name, TargetRow = targetRow, ReferenceRow = refRow });
            }
            else if (!foundExisting)
            {
                report.MissingReference.Add(item.Name);
            }
        }

        private static void FillCodeIfMissing(
            IXLWorksheet worksheet,
            IDictionary<string, string> rawCodeMap,
            PowerRow item,
            int targetRow,
            Stage1Report report)
        {
            if (!string.IsNullOrWhiteSpace(TextUtil.S(worksheet.Cell(targetRow, 2).GetFormattedString())))
            {
                return;
            }

            string code;
            if (rawCodeMap.TryGetValue(item.Key, out code))
            {
                worksheet.Cell(targetRow, 2).Value = code;
                report.CodeFilledFromRaw.Add(new RowMatchReport { Name = item.Name, TargetRow = targetRow, Code = code });
            }
            else
            {
                report.MissingCodes.Add(item.Name);
            }
        }

        private static void WriteMonthPower(IXLWorksheet worksheet, int targetStart, int targetRow, PowerRow item)
        {
            worksheet.Cell(targetRow, targetStart + 0).Value = item.Total;
            worksheet.Cell(targetRow, targetStart + 1).Value = item.Sharp;
            worksheet.Cell(targetRow, targetStart + 2).Value = item.Peak;
            worksheet.Cell(targetRow, targetStart + 3).Value = item.Flat;
            worksheet.Cell(targetRow, targetStart + 4).Value = item.Valley;
        }

        private static void FillPostWriteChecks(
            IXLWorksheet worksheet,
            IDictionary<string, int> ledgerMap,
            IList<PowerRow> powerRows,
            Stage1Report report)
        {
            foreach (var group in powerRows.GroupBy(row => row.Key).Where(group => group.Count() > 1))
            {
                report.DuplicateNamesInPowerFile[group.Key] = group.Count();
            }

            foreach (var item in powerRows)
            {
                var targetRow = ledgerMap[item.Key];
                if (string.IsNullOrWhiteSpace(TextUtil.S(worksheet.Cell(targetRow, 10).GetFormattedString())))
                {
                    report.MissingManualInfo.Add("第" + targetRow + "行 " + item.Name + " 缺少负责人/J列");
                }
            }

            report.MatchedRows = report.MatchedCustomers.Count;
            report.NewRows = report.NewCustomers.Count;
        }

        private static Dictionary<string, int> LedgerRowMap(IXLWorksheet worksheet)
        {
            var result = new Dictionary<string, int>();
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            for (var row = 4; row <= lastRow; row++)
            {
                var seq = TextUtil.N(worksheet.Cell(row, 1).Value);
                var name = TextUtil.S(worksheet.Cell(row, 3).GetFormattedString());
                if (seq > 0 && name.Length > 0)
                {
                    result[TextUtil.CustomerKey(name)] = row;
                }
            }
            return result;
        }

        private static void CopyMonthBlock(IXLWorksheet worksheet, int sourceMonth, int targetMonth)
        {
            var sourceStart = LedgerLayout.MonthStartColumn(sourceMonth);
            var targetStart = LedgerLayout.MonthStartColumn(targetMonth);
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            var sourceRange = worksheet.Range(1, sourceStart, lastRow, sourceStart + LedgerLayout.MonthBlockWidth - 1);
            sourceRange.CopyTo(worksheet.Cell(1, targetStart));
            worksheet.Cell(1, targetStart).Value = targetMonth + "月";

            for (var offset = 0; offset < LedgerLayout.MonthBlockWidth; offset++)
            {
                worksheet.Column(targetStart + offset).Width = worksheet.Column(sourceStart + offset).Width;
                if (worksheet.Column(sourceStart + offset).IsHidden)
                {
                    worksheet.Column(targetStart + offset).Hide();
                }
                else
                {
                    worksheet.Column(targetStart + offset).Unhide();
                }
            }
        }

        private static void CopyRowTemplate(IXLWorksheet worksheet, int styleTemplateRow, int targetRow, int month)
        {
            var maxCol = LedgerLayout.MonthStartColumn(month) + LedgerLayout.MonthBlockWidth - 1;
            worksheet.Row(styleTemplateRow).CopyTo(worksheet.Row(targetRow));
            for (var col = 1; col <= maxCol; col++)
            {
                if (col == 23 || col == 24 || col == 25 || col == 27 || col >= LedgerLayout.MonthStartColumn(month) + 5)
                {
                    continue;
                }
                worksheet.Cell(targetRow, col).Clear(XLClearOptions.Contents);
            }
        }

        private static void CopyBaseInfoFromReference(IXLWorksheet output, IXLWorksheet reference, int outputRow, int referenceRow)
        {
            for (var col = LedgerLayout.BaseStartColumn; col <= LedgerLayout.BaseEndColumn; col++)
            {
                reference.Cell(referenceRow, col).CopyTo(output.Cell(outputRow, col));
            }
        }
    }
}
