using NPOI.SS.UserModel;

namespace BlueColor.NPOIExtensions
{
    /// <summary>
    /// 行单元格
    /// </summary>
    public class RowCell
    {
        /// <summary>
        /// 单元格值
        /// </summary>
        public string CellValue { get; set; }

        /// <summary>
        /// 单元格样式
        /// </summary>
        public ICellStyle CellStyle { get; set; }
    }
}