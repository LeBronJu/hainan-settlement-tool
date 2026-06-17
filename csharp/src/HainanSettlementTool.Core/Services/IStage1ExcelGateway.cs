using System.Collections.Generic;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public interface IStage1ExcelGateway
    {
        List<PowerRow> ReadPowerRows(string powerPath);
        List<PowerRow> ReadRawPowerRows(string rawDetailPath);
        Dictionary<string, string> ReadCustomerCodes(string rawDetailPath);
        void WritePowerWorkbook(IEnumerable<PowerRow> rows, string outputPath);
        Stage1Report UpdateLedger(Stage1Options options);
    }
}
