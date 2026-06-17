using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal sealed class PowerWorkbookReader
    {
        public List<PowerRow> Read(string powerPath)
        {
            using (var workbook = new XLWorkbook(powerPath))
            {
                var worksheet = workbook.Worksheets.First();
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                var rows = new List<PowerRow>();

                for (var row = 2; row <= lastRow; row++)
                {
                    var name = TextUtil.S(worksheet.Cell(row, 1).GetValue<string>());
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    rows.Add(new PowerRow
                    {
                        SourceRow = row,
                        Name = name,
                        Key = TextUtil.CustomerKey(name),
                        Total = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 2)),
                        Sharp = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 3)),
                        Peak = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 4)),
                        Flat = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 5)),
                        Valley = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 6))
                    });
                }

                return rows;
            }
        }

        public void Write(IEnumerable<PowerRow> rows, string outputPath)
        {
            var grouped = rows
                .Where(row => !string.IsNullOrWhiteSpace(row.Key))
                .GroupBy(row => row.Key)
                .Select(group => new PowerRow
                {
                    Name = group.First().Name,
                    Key = group.Key,
                    Total = System.Math.Round(group.Sum(row => row.Total), 4),
                    Sharp = System.Math.Round(group.Sum(row => row.Sharp), 4),
                    Peak = System.Math.Round(group.Sum(row => row.Peak), 4),
                    Flat = System.Math.Round(group.Sum(row => row.Flat), 4),
                    Valley = System.Math.Round(group.Sum(row => row.Valley), 4)
                })
                .OrderBy(row => row.Name, System.StringComparer.CurrentCulture)
                .ToList();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");
                var headers = new[] { "零售用户名称", "总电量(I)", "尖段电量(L)", "峰段电量(P)", "平段电量(T)", "谷段电量(X)" };
                for (var col = 1; col <= headers.Length; col++)
                {
                    worksheet.Cell(1, col).Value = headers[col - 1];
                    worksheet.Cell(1, col).Style.Font.Bold = true;
                }

                for (var index = 0; index < grouped.Count; index++)
                {
                    var row = grouped[index];
                    var excelRow = index + 2;
                    worksheet.Cell(excelRow, 1).Value = row.Name;
                    worksheet.Cell(excelRow, 2).Value = row.Total;
                    worksheet.Cell(excelRow, 3).Value = row.Sharp;
                    worksheet.Cell(excelRow, 4).Value = row.Peak;
                    worksheet.Cell(excelRow, 5).Value = row.Flat;
                    worksheet.Cell(excelRow, 6).Value = row.Valley;
                }

                worksheet.Column(1).Width = 36;
                for (var col = 2; col <= 6; col++)
                {
                    worksheet.Column(col).Width = 14;
                }

                Directory.CreateDirectory(Path.GetDirectoryName(outputPath));
                workbook.SaveAs(outputPath);
            }
        }
    }
}
