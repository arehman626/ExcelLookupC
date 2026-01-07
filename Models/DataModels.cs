namespace ExcelLookupC.Models
{
    public class ComparisonResult
    {
        public List<Dictionary<string, object?>> MatchedRecords { get; set; } = new();
        public List<Dictionary<string, object?>> LeftOnlyRecords { get; set; } = new();
        public List<Dictionary<string, object?>> RightOnlyRecords { get; set; } = new();
        public string LeftFileName { get; set; } = string.Empty;
        public string RightFileName { get; set; } = string.Empty;
        public string LeftSheetName { get; set; } = string.Empty;
        public string RightSheetName { get; set; } = string.Empty;
        public List<string> ComparisonColumns { get; set; } = new();
        public DateTime ComparisonTimestamp { get; set; } = DateTime.Now;
    }

    public class ExcelFileInfo
    {
        public string FilePath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public List<string> SheetNames { get; set; } = new();
        public Dictionary<string, List<string>> SheetColumns { get; set; } = new();
    }

    public class SheetData
    {
        public string SheetName { get; set; } = string.Empty;
        public List<string> ColumnNames { get; set; } = new();
        public List<Dictionary<string, object?>> Records { get; set; } = new();
    }
}