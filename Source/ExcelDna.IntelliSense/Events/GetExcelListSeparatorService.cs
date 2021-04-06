#nullable enable


namespace ExcelDna.IntelliSense
{
    public class GetExcelListSeparatorService : IGetExcelListSeparator
    {
        public static GetExcelListSeparatorService Instance = new GetExcelListSeparatorService();

        public char ListSeparator => FormulaParser.ListSeparator;
    }
}
