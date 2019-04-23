using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Data;
using System.IO;

namespace BlueColor.NPOIExtensions.Samples
{
    public class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            CreateWorkbook_HSSFWorkbook();
        }

        /// <summary>
        /// HSSFWorkbook
        /// 创建工作簿
        /// </summary>
        public static void CreateWorkbook_HSSFWorkbook()
        {
            IWorkbook workbook = new HSSFWorkbook();
            var sheetA1 = workbook.CreateSheet("Sheet A1");
            var sheetA2 = workbook.CreateSheet("Sheet A2");

            sheetA1.AddMultipleRowsByPropertyCellConfig(
                propertyCellConfigs: Program.GetDemoPropertyCellConfigs(),
                dataTable: Program.GetDemoDataTable(),
                startRowNum: 0
                );

            //FileStream sw = File.Create("test.xlsx");
            //workbook.Write(sw);
            //sw.Close();

            var isOk = workbook.createFileToWrite("test.xlsx");
            Console.WriteLine($"创建完成：{isOk.ToString()}.");
        }

        /// <summary>
        /// 获取示例DataTable
        /// </summary>
        /// <returns></returns>
        public static DataTable GetDemoDataTable()
        {
            var table = new DataTable("Demo");

            DataColumn column;
            DataRow row;

            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.Caption = "Name";
            column.ColumnName = "Name";
            table.Columns.Add(column);

            column = new DataColumn();
            column.DataType = Type.GetType("System.Int32");
            column.Caption = "Age";
            column.ColumnName = "Age";
            column.DefaultValue = 21;
            table.Columns.Add(column);

            column = new DataColumn();
            column.DataType = Type.GetType("System.Decimal");
            column.Caption = "Growth";
            column.ColumnName = "Growth";
            column.DefaultValue = 0.21;
            table.Columns.Add(column);

            for (int i = 0; i <= 10; i++)
            {
                row = table.NewRow();
                row["Name"] = "AA";
                //row["Age"] = 21;
                row["Growth"] = 0.21;
                table.Rows.Add(row);
            }
            for (int i = 0; i <= 10; i++)
            {
                row = table.NewRow();
                row["Name"] = "BB";
                row["Age"] = 22;
                row["Growth"] = 0.21;
                table.Rows.Add(row);
            }

            return table;
        }

        /// <summary>
        /// 获取示例PropertyCellConfigs
        /// </summary>
        /// <returns></returns>
        public static PropertyCellConfigList GetDemoPropertyCellConfigs()
        {
            var configs = new PropertyCellConfigList { };

            configs.Add(new PropertyCellConfig
            {
                PropertyName = "Name",
                Title = "名称",
                ColumnIndex = 0,
                ConditionColumn = true,
                AllowMerge = true,
                IsIgnored = false,
                //DataFormat
            });
            configs.Add(new PropertyCellConfig
            {
                PropertyName = "Age",
                Title = "年龄",
                ColumnIndex = 1,
                ConditionColumn = false,
                AllowMerge = true,
                IsIgnored = false,
                DataFormat = "0"
            });
            configs.Add(new PropertyCellConfig
            {
                PropertyType = typeof(decimal),
                PropertyName = "Growth",
                Title = "成长",
                ColumnIndex = 2,
                ConditionColumn = false,
                AllowMerge = false,
                IsIgnored = false,
                DataFormat = "0.00%"
            });

            return configs;
        }
    }
}