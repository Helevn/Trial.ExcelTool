// See https://aka.ms/new-console-template for more information
using ConsoleApp1;


//读文件
var read = ExcelTool.ExcelServiceNPOI.ReadExcel<InfoExcelEntitys>("info.xlsx") ?? [];


read.Add(new InfoExcelEntitys { Code = Guid.NewGuid().ToString(), Name = Guid.NewGuid().ToString(), Age = Guid.NewGuid().ToString() });

//写文件
ExcelTool.ExcelServiceNPOI.WriteExcel(read, "info.xlsx");

Console.ReadKey();