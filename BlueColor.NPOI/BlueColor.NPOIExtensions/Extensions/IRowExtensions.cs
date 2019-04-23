using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace BlueColor.NPOIExtensions
{
    /// <summary>
    /// 设置单元格值-委托
    /// </summary>
    /// <param name="colIndex">列索引++</param>
    /// <param name="value">值</param>
    /// <param name="cellStyle">格式</param>
    public delegate void BcSetCellValueByCellStyleDelegate(ref int colIndex, object value, ICellStyle cellStyle = null);

    /// <summary>
    /// IRow 扩展
    /// </summary>
    public static class IRowExtensions
    {
        /// <summary>
        /// 获得单元格值
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="index">索引</param>
        /// <returns></returns>
        public static object GetCellValue(this IRow row, int index)
        {
            var cell = row.GetCell(index);
            if (cell == null)
            {
                return null;
            }

            if (cell.IsMergedCell)
            {
                // what can I do here?
            }

            switch (cell.CellType)
            {
                // This is a trick to get the correct value of the cell.
                // NumericCellValue will return a numeric value no matter the cell value is a date or a number.
                case CellType.Numeric:
                    return cell.ToString();

                case CellType.String:
                    return cell.StringCellValue;

                case CellType.Boolean:
                    return cell.BooleanCellValue;

                case CellType.Error:
                    return cell.ErrorCellValue;

                // how?
                case CellType.Formula:
                    return cell.ToString();

                case CellType.Blank:
                case CellType.Unknown:
                default:
                    return null;
            }
        }

        #region 设置单元格值

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="colIndex">索引</param>
        /// <param name="value">值</param>
        /// <param name="cellStyle">格式</param>
        public static void bcSetCellValueByCellStyle(this IRow row, ref int colIndex, object value, ICellStyle cellStyle = null)
        {
            throw new NotImplementedException("此功能暂未实现");
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="colIndex">索引</param>
        /// <param name="value">值</param>
        /// <param name="cellStyle">格式</param>
        public static void bcSetCellValueByCellStyle(this IRow row, ref int colIndex, bool value, ICellStyle cellStyle = null)
        {
            ICell cell = row.CreateCell(colIndex++);
            cell.SetCellValue(value);
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="colIndex">列索引++</param>
        /// <param name="value">值</param>
        /// <param name="cellStyle">格式</param>
        public static void bcSetCellValueByCellStyle(this IRow row, ref int colIndex, string value, ICellStyle cellStyle = null)
        {
            ICell cell = row.CreateCell(colIndex++);
            cell.SetCellValue(value);
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="colIndex">列索引++</param>
        /// <param name="value">值</param>
        /// <param name="cellStyle">格式</param>
        public static void bcSetCellValueByCellStyle(this IRow row, ref int colIndex, IRichTextString value, ICellStyle cellStyle = null)
        {
            ICell cell = row.CreateCell(colIndex++);
            cell.SetCellValue(value);
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="colIndex">列索引++</param>
        /// <param name="value">值</param>
        /// <param name="cellStyle">格式</param>
        public static void bcSetCellValueByCellStyle(this IRow row, ref int colIndex, DateTime value, ICellStyle cellStyle = null)
        {
            ICell cell = row.CreateCell(colIndex++);
            cell.SetCellValue(value);
            cell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="colIndex">列索引++</param>
        /// <param name="value">值</param>
        /// <param name="cellStyle">格式</param>
        public static void bcSetCellValueByCellStyle(this IRow row, ref int colIndex, double value, ICellStyle cellStyle = null)
        {
            ICell cell = row.CreateCell(colIndex++);
            cell.SetCellValue(value);
            cell.CellStyle = cellStyle;
        }

        #endregion 设置单元格值

        #region 设置单元格值-按配置

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="value">值</param>
        /// <param name="config"></param>
        /// <param name="cellStyles"></param>
        public static void bcSetCellValueByConfig(this IRow row, object value, PropertyCellConfig config, Dictionary<int, ICellStyle> cellStyles)
        {
            throw new NotImplementedException("此功能暂未实现");
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="value">值</param>
        /// <param name="config"></param>
        /// <param name="cellStyles"></param>
        public static void bcSetCellValueByConfig(this IRow row, bool value, PropertyCellConfig config, Dictionary<int, ICellStyle> cellStyles)
        {
            ICell cell = row.CreateCell(config.ColumnIndex);
            cell.SetCellValue(value);
            cell.CellStyle = cellStyles[config.ColumnIndex];
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="value">值</param>
        /// <param name="config"></param>
        /// <param name="cellStyles"></param>
        public static void bcSetCellValueByConfig(this IRow row, string value, PropertyCellConfig config, Dictionary<int, ICellStyle> cellStyles)
        {
            ICell cell = row.CreateCell(config.ColumnIndex);
            cell.SetCellValue(value);
            cell.CellStyle = cellStyles[config.ColumnIndex];
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="value">值</param>
        /// <param name="config"></param>
        /// <param name="cellStyles"></param>
        public static void bcSetCellValueByConfig(this IRow row, IRichTextString value, PropertyCellConfig config, Dictionary<int, ICellStyle> cellStyles)
        {
            ICell cell = row.CreateCell(config.ColumnIndex);
            cell.SetCellValue(value);
            cell.CellStyle = cellStyles[config.ColumnIndex];
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="value">值</param>
        /// <param name="config"></param>
        /// <param name="cellStyles"></param>
        public static void bcSetCellValueByConfig(this IRow row, DateTime value, PropertyCellConfig config, Dictionary<int, ICellStyle> cellStyles)
        {
            ICell cell = row.CreateCell(config.ColumnIndex);
            cell.SetCellValue(value);
            cell.CellStyle = cellStyles[config.ColumnIndex];
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="value">值</param>
        /// <param name="config"></param>
        /// <param name="cellStyles"></param>
        public static void bcSetCellValueByConfig(this IRow row, double value, PropertyCellConfig config, Dictionary<int, ICellStyle> cellStyles)
        {
            ICell cell = row.CreateCell(config.ColumnIndex);
            cell.SetCellValue(value);
            cell.CellStyle = cellStyles[config.ColumnIndex];
        }

        #endregion 设置单元格值-按配置

    }
}