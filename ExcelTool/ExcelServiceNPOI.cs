using Microsoft.VisualBasic;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    public class ExcelServiceNPOI
    {
        public static IList<T> ReadExcel<T>(string file)
        {
            var list = new List<T>();
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();

            using (FileStream fs = new(file, FileMode.Open, FileAccess.Read))
            {
                if (fileExt == ".xlsx")
                    workbook = new XSSFWorkbook(fs);
                else
                    throw new Exception($"未能识别格式的文件【{fileExt}】");

                ISheet sheet = workbook.GetSheetAt(0) ?? throw new Exception($"未能找到有效的sheet的数据");
                //取特性
                var props = typeof(T).GetProperties();
                var attrrs = props.Select(s => s.GetCustomAttribute<ColumnNameAttribute>()).ToArray();
                //先找到对应特性名称的数据行
                var startRow = -1;
                for (int i = 0; i <= sheet.LastRowNum; i++)
                {
                    var cell = sheet.GetRow(i)?.GetCell(0)?.ToString() ?? "";//取该行的第0列为特征点
                    if (cell == attrrs[0].Name)
                    {
                        startRow = i;
                        break;
                    }
                }
                if (startRow == -1)
                    throw new Exception("未找到对应特姓名称的开始数据行");

                //取表头  
                IRow header = sheet.GetRow(startRow);
                for (int i = 0; i < attrrs.Length; i++)
                {
                    var headerName = GetValueType(header.GetCell(i)).ToString() ?? "";//校验表头名称,默认数据是从第0列开始的
                    if (attrrs[i] != null && headerName != attrrs[i].Name)
                    {
                        throw new Exception($"Excel表头第【{i + 1}】列识别错误:无法识别的表头名称:{headerName}");
                    }
                }
                //取数据
                for (int i = startRow + 1; i <= sheet.LastRowNum; i++)
                {
                    var p = Activator.CreateInstance<T>();
                    var currentRow = sheet.GetRow(i);
                    if (currentRow == null)
                        break;//当发现某一行所有数据为空时，代表该组数据已经读取完成
                    if (currentRow.Cells.Where(s => s.ColumnIndex >= 0 && s.ColumnIndex < attrrs.Length)
                        .All(s => string.IsNullOrEmpty(GetValueType(s)?.ToString() ?? "")))//当发现某一行所有数据为空时跳过该行，直接读取
                        continue;
                    for (int j = 0; j < props.Length; j++)
                    {
                        var cell = GetValueType(currentRow.GetCell(j))?.ToString() ?? "";
                        var prop = props[j];
                        prop.SetValue(p, cell);
                    }
                    list.Add(p);
                }
            }
            object GetValueType(ICell cell)
            {
                if (cell == null)
                    return null;
                return cell.CellType switch
                {
                    CellType.Blank => null,
                    CellType.Boolean => cell.BooleanCellValue,
                    CellType.Numeric => cell.NumericCellValue,
                    CellType.String => cell.StringCellValue,
                    CellType.Error => cell.ErrorCellValue,
                    CellType.Formula => cell.RichStringCellValue,
                    _ => "=" + cell.CellFormula,
                };
            }
            return list;
        }
        public static string WriteExcel<T>(IList<T> data, string file)
        {
            IWorkbook workbook = new XSSFWorkbook();
            string fileExt = Path.GetExtension(file).ToLower();

            var fileinfo = new FileInfo(file);
            if (fileinfo.Exists)
                fileinfo.Delete();
            ISheet sheet = workbook.CreateSheet(typeof(T).Name);
            //取特性
            var props = typeof(T).GetProperties();

            var attrrs = props.Select(s => s.GetCustomAttribute<ColumnNameAttribute>()).ToArray();

            //写表头
            var header = sheet.CreateRow(0);
            foreach (var item in attrrs)
            {
                var index = Array.IndexOf(attrrs, item);
                var cell = header.CreateCell(index);
                cell.SetCellValue(item.Name);
            }

            //写数据  
            for (int i = 1; i <= data.Count(); i++)
            {
                var d = data[i - 1];
                IRow rowData = sheet.CreateRow(i);
                foreach (var item in attrrs)
                {
                    var index = Array.IndexOf(attrrs, item);
                    var cell = rowData.CreateCell(index);
                    cell.SetCellValue(props[index].GetValue(d)?.ToString() ?? "");
                }
            }
            //excel写入
            using (FileStream fs_write = new FileStream(file, FileMode.CreateNew, FileAccess.Write))
            {
                workbook.Write(fs_write);
            }
            return Path.GetDirectoryName(file);
        }
    }
}
