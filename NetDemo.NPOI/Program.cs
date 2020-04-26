using System;

namespace NetDemo.NPOI
{
    class Program
    {
        static void Main(string[] args)
        {
            var dt = ExcelHelper.ExcelToDataTable(@"D:\文档\20200423销售指标线上化\业绩指标模板-终版.xlsx", false);
            ExcelHelper.DataTableToExcel(dt, @"D:\文档\20200423销售指标线上化\业绩指标模板-终版1.xlsx");
            Console.WriteLine("Hello World!");
            Console.ReadLine();
        }
    }
}
