using ExcelTool;

namespace ConsoleApp1
{
    public class InfoExcelEntitys
    {
        [ColumnName("Code")]
        public string Code { get; set; } = "";
        [ColumnName("Name")]
        public string Name { get; set; } = "";
        [ColumnName("Age")]
        public string Age { get; set; } = "";
    }
}
