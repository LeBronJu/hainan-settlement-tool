using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal sealed class CustomerCodeReader
    {
        public Dictionary<string, string> Read(string rawDetailPath)
        {
            var result = new Dictionary<string, string>();
            if (string.IsNullOrWhiteSpace(rawDetailPath) || !File.Exists(rawDetailPath))
            {
                return result;
            }

            if (!string.Equals(Path.GetExtension(rawDetailPath), ".xlsx", System.StringComparison.OrdinalIgnoreCase))
            {
                return result;
            }

            using (var workbook = new XLWorkbook(rawDetailPath))
            {
                var sheetNames = new[] { "零售主体电量", "零售户号电量" }
                    .Where(name => workbook.Worksheets.Any(ws => ws.Name == name))
                    .ToList();
                if (sheetNames.Count == 0)
                {
                    sheetNames.Add(workbook.Worksheets.First().Name);
                }

                foreach (var sheetName in sheetNames)
                {
                    var worksheet = workbook.Worksheet(sheetName);
                    var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                    for (var row = 4; row <= lastRow; row++)
                    {
                        var code = TextUtil.S(worksheet.Cell(row, 3).GetFormattedString());
                        var name = TextUtil.S(worksheet.Cell(row, 4).GetFormattedString());
                        var key = TextUtil.CustomerKey(name);
                        if (key.Length > 0 && code.Length > 0 && !result.ContainsKey(key))
                        {
                            result[key] = code;
                        }
                    }
                }
            }

            return result;
        }
    }
}
