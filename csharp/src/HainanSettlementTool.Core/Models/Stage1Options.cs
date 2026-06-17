namespace HainanSettlementTool.Core.Models
{
    public sealed class Stage1Options
    {
        public int Month { get; set; }
        public string BaseLedgerPath { get; set; }
        public string PowerPath { get; set; }
        public string RawDetailPath { get; set; }
        public string ReferenceLedgerPath { get; set; }
        public string OutputDirectory { get; set; }
        public string OutputLedgerName { get; set; }
        public bool CopyReferenceExisting { get; set; }
    }
}
