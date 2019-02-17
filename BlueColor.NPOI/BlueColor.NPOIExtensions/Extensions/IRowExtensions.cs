using NPOI.SS.UserModel;

namespace BlueColor.NPOIExtensions
{
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
    }
}