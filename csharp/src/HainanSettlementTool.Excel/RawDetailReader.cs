using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal sealed class RawDetailReader
    {
        public List<PowerRow> Read(string rawDetailPath)
        {
            var extension = Path.GetExtension(rawDetailPath).ToLowerInvariant();
            if (extension == ".xlsx")
            {
                return ReadXlsx(rawDetailPath);
            }

            if (extension == ".csv")
            {
                return ReadCsv(rawDetailPath);
            }

            throw new NotSupportedException("C# 第一版暂不直接清洗 .xls，请先另存为 .xlsx 或使用已清洗电量表。");
        }

        private static List<PowerRow> ReadXlsx(string path)
        {
            using (var workbook = new XLWorkbook(path))
            {
                var worksheet = workbook.Worksheets.First();
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                var rows = new List<PowerRow>();
                for (var row = 4; row <= lastRow; row++)
                {
                    var name = TextUtil.S(worksheet.Cell(row, 4).GetFormattedString());
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    rows.Add(new PowerRow
                    {
                        SourceRow = row,
                        Name = name,
                        Key = TextUtil.CustomerKey(name),
                        Total = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 9)),
                        Sharp = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 12)),
                        Peak = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 16)),
                        Flat = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 20)),
                        Valley = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 24))
                    });
                }

                return rows;
            }
        }

        private static List<PowerRow> ReadCsv(string path)
        {
            var rows = new List<PowerRow>();
            var lines = File.ReadAllLines(path, DetectCsvEncoding(path));
            for (var index = 3; index < lines.Length; index++)
            {
                var cols = SplitCsvLine(lines[index]);
                if (cols.Count <= 23 || string.IsNullOrWhiteSpace(cols[3]))
                {
                    continue;
                }

                var name = TextUtil.S(cols[3]);
                rows.Add(new PowerRow
                {
                    SourceRow = index + 1,
                    Name = name,
                    Key = TextUtil.CustomerKey(name),
                    Total = TextUtil.N(cols[8]),
                    Sharp = TextUtil.N(cols[11]),
                    Peak = TextUtil.N(cols[15]),
                    Flat = TextUtil.N(cols[19]),
                    Valley = TextUtil.N(cols[23])
                });
            }

            return rows;
        }

        private static Encoding DetectCsvEncoding(string path)
        {
            var bytes = File.ReadAllBytes(path);
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            {
                return new UTF8Encoding(true);
            }

            return Encoding.GetEncoding("GB18030");
        }

        private static List<string> SplitCsvLine(string line)
        {
            var result = new List<string>();
            var sb = new StringBuilder();
            var quoted = false;
            for (var i = 0; i < line.Length; i++)
            {
                var ch = line[i];
                if (ch == '"')
                {
                    if (quoted && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        quoted = !quoted;
                    }
                }
                else if (ch == ',' && !quoted)
                {
                    result.Add(sb.ToString());
                    sb.Length = 0;
                }
                else
                {
                    sb.Append(ch);
                }
            }

            result.Add(sb.ToString());
            return result;
        }
    }
}
