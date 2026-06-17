using System.Collections.Generic;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    public sealed class ClosedXmlStage1ExcelGateway : IStage1ExcelGateway
    {
        private readonly PowerWorkbookReader _powerWorkbookReader = new PowerWorkbookReader();
        private readonly RawDetailReader _rawDetailReader = new RawDetailReader();
        private readonly CustomerCodeReader _customerCodeReader = new CustomerCodeReader();
        private readonly LedgerStage1Updater _ledgerUpdater = new LedgerStage1Updater();

        public List<PowerRow> ReadPowerRows(string powerPath)
        {
            return _powerWorkbookReader.Read(powerPath);
        }

        public List<PowerRow> ReadRawPowerRows(string rawDetailPath)
        {
            return _rawDetailReader.Read(rawDetailPath);
        }

        public Dictionary<string, string> ReadCustomerCodes(string rawDetailPath)
        {
            return _customerCodeReader.Read(rawDetailPath);
        }

        public void WritePowerWorkbook(IEnumerable<PowerRow> rows, string outputPath)
        {
            _powerWorkbookReader.Write(rows, outputPath);
        }

        public Stage1Report UpdateLedger(Stage1Options options)
        {
            return _ledgerUpdater.Update(options, this);
        }
    }
}
