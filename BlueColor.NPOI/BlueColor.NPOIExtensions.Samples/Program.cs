using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace BlueColor.NPOIExtensions.Samples
{
    public class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            //CreateWorkbook_HSSFWorkbook();
            //CreateWorkbook_HSSFWorkbook2();
            CreateWorkbook_HSSFWorkbook3();
        }

        #region 示例1

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

        #endregion

        #region 示例2

        /// <summary>
        /// HSSFWorkbook
        /// 创建工作簿
        /// </summary>
        public static void CreateWorkbook_HSSFWorkbook2()
        {
            IWorkbook workbook = new HSSFWorkbook();
            var sheet1 = workbook.CreateSheet("Sheet1");

            sheet1.AddMultipleRowsContentByList(null, genTestAList(), handleTestAProperty);

            var isOk = workbook.createFileToWrite("test2.xlsx");
            Console.WriteLine($"创建完成：{isOk.ToString()}.");
        }

        /// <summary>
        /// HSSFWorkbook
        /// 创建工作簿
        /// </summary>
        public static void CreateWorkbook_HSSFWorkbook3()
        {
            IWorkbook workbook = new HSSFWorkbook();
            var sheet1 = workbook.CreateSheet("Sheet1");

            sheet1.AddMultipleRowsByPropertyCellConfigByList(Program.GetDemoPropertyCellConfigs(), genTestAList(), handleTestAPropertyByPropertyCellConfigList);

            var isOk = workbook.createFileToWrite("test3.xlsx");
            Console.WriteLine($"创建完成：{isOk.ToString()}.");
        }

        /// <summary>
        /// 生成 TestAList
        /// </summary>
        /// <returns></returns>
        public static List<TestA> genTestAList()
        {
            var testAs = new List<TestA>() { };

            testAs.Add(new TestA() { Name = "AA", Age = 21, Growth = 0.21m });
            testAs.Add(new TestA() { Name = "AA", Age = 21, Growth = 0.21m });
            testAs.Add(new TestA() { Name = "AA", Age = 21, Growth = 0.21m });
            testAs.Add(new TestA() { Name = "AA", Age = 21, Growth = 0.21m });
            testAs.Add(new TestA() { Name = "AA", Age = 21, Growth = 0.21m });

            testAs.Add(new TestA() { Name = "BB", Age = 22, Growth = 0.22m });
            testAs.Add(new TestA() { Name = "BB", Age = 22, Growth = 0.22m });
            testAs.Add(new TestA() { Name = "BB", Age = 22, Growth = 0.22m });
            testAs.Add(new TestA() { Name = "BB", Age = 22, Growth = 0.22m });
            testAs.Add(new TestA() { Name = "BB", Age = 22, Growth = 0.22m });

            return testAs;
        }

        /// <summary>
        /// 处理 TestA 属性映射
        /// </summary>
        /// <param name="row"></param>
        /// <param name="startColNum"></param>
        /// <param name="testA"></param>
        /// <param name="cellStyle"></param>
        public static void handleTestAProperty(IRow row, int startColNum, TestA testA, ICellStyle cellStyle)
        {
            row.bcSetCellValueByCellStyle(ref startColNum, testA.Name, cellStyle);
            row.bcSetCellValueByCellStyle(ref startColNum, testA.Age, cellStyle);
            row.bcSetCellValueByCellStyle(ref startColNum, testA.Growth.ToString(), cellStyle);
        }

        /// <summary>
        /// 处理 TestA 属性映射
        /// </summary>
        /// <param name="row"></param>
        /// <param name="testA"></param>
        /// <param name="propertyCellConfigs"></param>
        /// <param name="cellStyles"></param>
        public static void handleTestAPropertyByPropertyCellConfigList(IRow row, TestA testA, PropertyCellConfigList propertyCellConfigs, Dictionary<int, ICellStyle> cellStyles)
        {
            row.bcSetCellValueByConfig(testA.Name, propertyCellConfigs.GetConfigByName("Name"), cellStyles);
            row.bcSetCellValueByConfig(testA.Age, propertyCellConfigs.GetConfigByName("Age"), cellStyles);
            row.bcSetCellValueByConfig(Convert.ToDouble(testA.Growth), propertyCellConfigs.GetConfigByName("Growth"), cellStyles);
        }

        #endregion
    }
}