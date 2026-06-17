namespace HainanSettlementTool.Core.Services
{
    public static class LedgerLayout
    {
        public const string MainSheetName = "海南2026年售电结算台账";
        public const int FirstMonthStartColumn = 32; // AF
        public const int MonthBlockWidth = 26;
        public const int BaseStartColumn = 2;
        public const int BaseEndColumn = 28;

        public static int MonthStartColumn(int month)
        {
            return FirstMonthStartColumn + (month - 1) * MonthBlockWidth;
        }
    }
}
