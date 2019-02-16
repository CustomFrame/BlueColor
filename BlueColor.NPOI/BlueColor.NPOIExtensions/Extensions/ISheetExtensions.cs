using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace BlueColor.NPOIExtensions.Extensions
{
    /// <summary>
    /// ISheet 扩展
    /// </summary>
    public static class ISheetExtensions
    {
        /// <summary>
        /// 添加单行标题
        /// (单行标题无合并单元格情况)
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="columnNames">列名-List</param>
        /// <param name="startRowNum">开始行号</param>
        /// <param name="cellStyle">单元格式</param>
        /// <returns></returns>
        public static void AddSingleRowHeader(this ISheet sheet, List<string> columnNames, int startRowNum = 0, ICellStyle cellStyle = null)
        {
            //单元格样式为null,置为默认
            cellStyle = cellStyle ?? sheet.Workbook.genHeaderCellStyle();

            IRow row = sheet.CreateRow(startRowNum);

            for (int i = 0; i < columnNames.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(columnNames[i]);
                cell.CellStyle = cellStyle;
            }
        }

        /// <summary>
        /// 添加单行内容
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="singleRowCells">单行单元格-List</param>
        /// <param name="startRowNum">开始行号</param>
        /// <param name="cellStyle">单元格式</param>
        /// <returns></returns>
        public static void AddSingleRowContent(this ISheet sheet, List<string> singleRowCells, int startRowNum = 1, ICellStyle cellStyle = null)
        {
            //单元格样式为null,置为默认
            cellStyle = cellStyle ?? sheet.Workbook.genContentCellStyle();

            IRow row = sheet.CreateRow(startRowNum);

            for (int i = 0; i < singleRowCells.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(singleRowCells[i]);
                cell.CellStyle = cellStyle;
            }
        }

        /// <summary>
        /// 添加单行内容
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="singleRowCells">单行单元格-List</param>
        /// <param name="startRowNum">开始行号</param>
        /// <param name="cellStyle">单元格式</param>
        /// <returns></returns>
        public static void AddSingleRowContentWithCellStyle(this ISheet sheet, List<RowCell> singleRowCells, int startRowNum = 1)
        {
            //单元格样式为null,置为默认
            var defaultCellStyle = sheet.Workbook.genContentCellStyle();

            IRow row = sheet.CreateRow(startRowNum);

            for (int i = 0; i < singleRowCells.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(singleRowCells[i].CellValue);
                cell.CellStyle = singleRowCells[i].CellStyle ?? defaultCellStyle;
            }
        }

        /// <summary>
        /// 添加多行内容
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="columnNames">列名-List</param>
        /// <param name="dataTable">数据表</param>
        /// <param name="startRowNum">开始行号</param>
        /// <param name="cellStyle">单元格式</param>
        /// <returns></returns>
        public static void AddMultipleRowsContent(this ISheet sheet, List<string> columnNames, DataTable dataTable, int startRowNum = 1, ICellStyle cellStyle = null)
        {
            //单元格样式为null,置为默认
            cellStyle = cellStyle ?? sheet.Workbook.genContentCellStyle();

            foreach (DataRow dataRow in dataTable.Rows)
            {
                IRow cells = sheet.CreateRow(startRowNum++);

                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    ICell cell = cells.CreateCell(i);
                    cell.SetCellValue(dataRow[i].ToString().Trim());
                    cell.CellStyle = cellStyle;
                }
            }

            IRow row = sheet.CreateRow(startRowNum);
        }

        /// <summary>
        /// 添加标题行-输出列样式
        /// (按列-属性单元格配置)
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="propertyCellConfigs">属性单元格配置-List</param>
        /// <param name="titleRowNum">标题行号</param>
        /// <param name="cellStyles">输出列样式</param>
        /// <returns></returns>
        public static void AddTitleRowByPropertyCellConfig(this ISheet sheet, PropertyCellConfigList propertyCellConfigs, int titleRowNum, out Dictionary<int, ICellStyle> cellStyles)
        {
            propertyCellConfigs.ValidateError();

            cellStyles = new Dictionary<int, ICellStyle>();

            var titleStyle = sheet.Workbook.genHeaderCellStyle();
            var titleRow = sheet.CreateRow(titleRowNum);
            foreach (var config in propertyCellConfigs)
            {
                if (config.IsIgnored)
                    continue;//

                if (!string.IsNullOrWhiteSpace(config.DataFormat))
                {
                    var style = sheet.Workbook.genContentCellStyle();

                    var dataFormat = sheet.Workbook.CreateDataFormat();
                    style.DataFormat = dataFormat.GetFormat(config.DataFormat);
                    //style.DataFormat = HSSFDataFormat.GetBuiltinFormat(config.DataFormat);
                    cellStyles[config.ColumnIndex] = style;
                }

                var titleCell = titleRow.CreateCell(config.ColumnIndex);
                titleCell.CellStyle = titleStyle;
                titleCell.SetCellValue(config.Title);
            }
        }

        /// <summary>
        /// 添加多行
        /// (按列-属性单元格配置)
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="propertyCellConfigs">属性单元格配置-List</param>
        /// <param name="dataTable">数据表</param>
        /// <param name="startRowNum">开始行号</param>
        /// <returns></returns>
        public static void AddMultipleRowsByPropertyCellConfig(this ISheet sheet, PropertyCellConfigList propertyCellConfigs, DataTable dataTable, int startRowNum = 0)
        {// 性能与功能取舍（合久必分；分久必合；）、拆分导致if语句多次判断性能问题、合并导致复杂度升高；
            propertyCellConfigs.ValidateError();

            sheet.AddTitleRowByPropertyCellConfig(propertyCellConfigs, startRowNum, out var cellStyles);

            // 内容开始行号
            var contentStartRowNum = startRowNum + 1;
            var rowIndex = contentStartRowNum;
            foreach (DataRow dataRow in dataTable.Rows)
            {
                var row = sheet.CreateRow(rowIndex);

                foreach (var config in propertyCellConfigs)
                {
                    if (config.IsIgnored)
                        continue;

                    var value = dataRow[config.PropertyName];
                    if (value == null)
                        continue;

                    var cell = row.CreateCell(config.ColumnIndex);
                    if (cellStyles.TryGetValue(config.ColumnIndex, out var cellStyle))
                    {
                        cell.CellStyle = cellStyle;
                        cell.SetCellValue(value.ToString());
                    }
                    else
                    {
                        cell.SetCellValue(value.ToString());
                    }
                }

                rowIndex++;
            }

            sheet.MergeCellByPropertyCellConfig(propertyCellConfigs, contentStartRowNum, rowIndex);

            for (int i = 0; i < propertyCellConfigs.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }
        }

        /// <summary>
        /// 合并单元格
        /// (按列-属性单元格配置)
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="propertyCellConfigs">属性单元格配置-List</param>
        /// <param name="firstRow">开始行号</param>
        /// <param name="lastRow">结束行号</param>
        /// <returns></returns>
        public static void MergeCellByPropertyCellConfig(this ISheet sheet, PropertyCellConfigList propertyCellConfigs, int firstRow, int lastRow)
        {
            propertyCellConfigs.ValidateError();

            var configConditionColumn = propertyCellConfigs.SingleOrDefault(p => p.ConditionColumn);
            var configAllowMergeColumns = propertyCellConfigs.Where(p => p.AllowMerge);

            if (configConditionColumn != null)
            {
                var vStyle = sheet.Workbook.genContentCellStyle();
                vStyle.VerticalAlignment = VerticalAlignment.Center;

                object previous = null;
                int rowspan = 0;
                int rowIndex = firstRow;
                for (rowIndex = firstRow; rowIndex < lastRow; rowIndex++)
                {
                    var value = sheet.GetRow(rowIndex).GetCellValue(configConditionColumn.ColumnIndex);
                    if (object.Equals(previous, value) && value != null)
                    {
                        rowspan++;
                    }
                    else
                    {
                        if (rowspan > 1)
                        {
                            sheet.GetRow(rowIndex - rowspan).Cells[configConditionColumn.ColumnIndex].CellStyle = vStyle;
                            foreach (var config in configAllowMergeColumns)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(rowIndex - rowspan, rowIndex - 1, config.ColumnIndex, config.ColumnIndex));
                            }
                        }
                        rowspan = 1;
                        previous = value;
                    }
                }

                // 合并所有行
                if (rowspan > 1)
                {
                    sheet.GetRow(rowIndex - rowspan).Cells[configConditionColumn.ColumnIndex].CellStyle = vStyle;
                    foreach (var config in configAllowMergeColumns)
                    {
                        sheet.AddMergedRegion(new CellRangeAddress(rowIndex - rowspan, rowIndex - 1, config.ColumnIndex, config.ColumnIndex));
                    }
                }
            }
        }

        /// <summary>
        /// ISheet
        /// 添加合计行
        /// </summary>
        /// <param name="sheet"></param>
        public static void AddFooter(this ISheet sheet)
        {
            throw new NotImplementedException();
        }
    }
}