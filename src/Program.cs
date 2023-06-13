using ExcelKit.Core.Attributes;
using ExcelKit.Core.ExcelRead;
using ExcelKit.Core.Helpers;
using ExcelKit.Core.Infrastructure.Factorys;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace BigDataExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("测试ExcelKit导出");
            DynamicExport();
        }
        /// <summary>
        /// 并发线程导出示例
        /// </summary>
        private static void TaskExport()
        {
            string filename = "测试导出文件xlsx";
            if (IsFileInUse(filename + ".xlsx"))
            {
                Console.WriteLine("文件正在占用，请先关闭文件");
            }
            else
            {
                using (var context = ContextFactory.GetWriteContext(filename))
                {
                    // 创建第一个sheet 采用了并发多Sheet导出，一个线程一个Sheet，非必须
                    Parallel.For(1, 4, index =>
                    {
                        var sheet = context.CrateSheet<Person>($"Sheet1");
                        for (int i = 0; i < 1000; i++)
                        {
                            sheet.AppendData($"Sheet1", new Person { UserName = $"1-{i}-小明", Address = $"1-{i}-苏州", Age = 30, Hobby = "读书", Pwd = "123456" });
                        }
                    });
                }
            }       
        
        }
        /// <summary>
        /// 基于实体类导出，需要在实体类配置ExcelKit的注解
        /// </summary>
        private static void Export()
        {
            string filename = "测试导出文件";
            if (IsFileInUse(filename + ".xlsx"))
            {
                Console.WriteLine("文件正在占用，请先关闭文件");
            }
            else
            {
                using (var context = ContextFactory.GetWriteContext("测试导出文件"))
                {
                    // 创建第一个sheet
                    var sheet = context.CrateSheet<Person>($"Sheet1");
                    for (int i = 0; i < 1000; i++)
                    {
                        sheet.AppendData($"Sheet1", new Person { UserName = $"1-{i}-小明", Address = $"1-{i}-苏州", Age = 30, Hobby = "读书", Pwd = "123456" });
                    }

                    // 创建第二个sheet
                    var sheet2 = context.CrateSheet<Person>($"Sheet2");
                    for (int i = 0; i < 1500; i++)
                    {
                        sheet2.AppendData($"Sheet2", new Person { UserName = $"2-{i}-小明", Address = $"2-{i}-苏州", Age = 30, Hobby = "读书", Pwd = "123456" });
                    }
                    string filePath = context.Save();
                    Console.WriteLine($"文件路径：{filePath}");
                }
            }
        }
        /// <summary>
        /// 动态指定列导出
        /// </summary>
        private static void DynamicExport()
        {
            string filename = "动态测试导出文件";
            if (IsFileInUse(filename+ ".xlsx"))
            {
                Console.WriteLine("文件正在占用，请先关闭文件");
            }
            else
            {
                using (var context = ContextFactory.GetWriteContext(filename))
                {
                    // 定义导出列属性
                    List<ExcelKitAttribute> excelKitAttributes = new List<ExcelKitAttribute>();
                    excelKitAttributes.Add(new ExcelKitAttribute
                    {
                        Code = "UserName",
                        Desc = "用户名",
                        Width = 30,
                        Sort = 10
                    });
                    excelKitAttributes.Add(new ExcelKitAttribute
                    {
                        Code = "Pwd",
                        Desc = "密码",
                        Width = 20,
                        Sort = 20
                    });

                    excelKitAttributes.Add(new ExcelKitAttribute
                    {
                        Code = "Address",
                        Desc = "住址",
                        Width = 50,
                        Sort = 30
                    });

                    // 创建第一个自定义列属性的sheet
                    var sheet = context.CrateSheet($"Sheet1", excelKitAttributes);
                    // 循环插入行记录
                    for (int i = 0; i < 600; i++)
                    {
                        sheet.AppendData($"Sheet1", new Dictionary<string, object> {
                        {"UserName",$"小明-{i}"}
                        ,{"Pwd","123456" }
                        ,{"Address",$"苏州-{i}"}
                    });
                    }
                    // 保存文件
                    string filePath = context.Save();
                    Console.WriteLine($"动态导出文件路径：{filePath}");
                }
            }
        }

        /// <summary>
        /// 读取表头
        /// </summary>
        private static void ReadHeaders()
        {
            //sheetIndex为Sheet索引(从1开始)，rowLine为行号(从1开始) 一般第一行表示列头，具体根据实际情况确定
            var headers = LiteDataHelper.ReadOneRow(filePath: "动态测试导出文件.xlsx", sheetIndex: 1, rowLine: 1);
            Console.WriteLine($"表头为：{string.Join("  ", headers)}");
        }
        /// <summary>
        /// 读取行数据和行数
        /// </summary>
        private static void ReadExcelDatas()
        {
            var context = ContextFactory.GetReadContext();
            // 读取行数 包含列头
            var count = context.ReadSheetRowsCount("动态测试导出文件.xlsx", new ReadSheetRowsCountOptions { });
            // 读取行数据 
            StringBuilder sb = new StringBuilder();
            context.ReadRows("动态测试导出文件.xlsx", new ReadRowsOptions()
            {
                RowData = rowdata =>
                {
                    sb.Append(JsonConvert.SerializeObject(rowdata) + "\n");
                }
            });
            Console.WriteLine($"读取的数据为:{sb.ToString()}");
        }

        /// <summary>
        /// 判断文件是否被占用
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static bool IsFileInUse(string fileName)
        {
            bool inUse = true;

            FileStream fs = null;
            try
            {
                fs = new FileStream(fileName, FileMode.Open,FileAccess.Read,FileShare.None);
                inUse = false;
            }
            catch
            {
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }                    
            }
            return inUse;//true表示正在使用,false没有使用
        }

    }
}
