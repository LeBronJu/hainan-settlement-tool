using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public sealed class Stage1Service
    {
        private readonly IStage1ExcelGateway _excel;

        public Stage1Service(IStage1ExcelGateway excel)
        {
            _excel = excel;
        }

        public Stage1Report Run(Stage1Options options, Action<string> log)
        {
            Validate(options);
            Directory.CreateDirectory(options.OutputDirectory);

            if (!File.Exists(options.PowerPath))
            {
                if (string.IsNullOrWhiteSpace(options.RawDetailPath) || !File.Exists(options.RawDetailPath))
                {
                    throw new FileNotFoundException("找不到电量处理表，也没有可清洗的原始零售侧明细。", options.PowerPath);
                }

                log?.Invoke("正在清洗原始零售侧明细，生成电量处理表。");
                var rawRows = _excel.ReadRawPowerRows(options.RawDetailPath);
                _excel.WritePowerWorkbook(rawRows, options.PowerPath);
            }

            log?.Invoke("正在把电量、新增客户名称和户号导入台账。");
            return _excel.UpdateLedger(options);
        }

        private static void Validate(Stage1Options options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (options.Month <= 1)
            {
                throw new ArgumentException("阶段1要求月份大于1。");
            }

            RequireExistingFile(options.BaseLedgerPath, "基础台账");
            if (string.IsNullOrWhiteSpace(options.PowerPath))
            {
                throw new ArgumentException("请选择电量处理表，或选择原始零售侧明细让程序生成。");
            }

            if (string.IsNullOrWhiteSpace(options.OutputDirectory))
            {
                throw new ArgumentException("请选择输出文件夹。");
            }

            if (!string.IsNullOrWhiteSpace(options.ReferenceLedgerPath))
            {
                RequireExistingFile(options.ReferenceLedgerPath, "参考台账");
            }

            if (!string.IsNullOrWhiteSpace(options.RawDetailPath) && !File.Exists(options.RawDetailPath))
            {
                throw new FileNotFoundException("原始零售侧明细不存在。", options.RawDetailPath);
            }
        }

        private static void RequireExistingFile(string path, string label)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("请选择" + label + "。");
            }

            if (Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
            {
                throw new ArgumentException(label + "选到了 Excel 临时文件，请选择正式文件。");
            }

            if (!File.Exists(path))
            {
                throw new FileNotFoundException(label + "不存在。", path);
            }
        }
    }
}
