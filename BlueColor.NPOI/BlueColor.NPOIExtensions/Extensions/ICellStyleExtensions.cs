using NPOI.SS.UserModel;

namespace BlueColor.NPOIExtensions
{
    /// <summary>
    /// ICellStyle 扩展
    /// </summary>
    public static class ICellStyleExtensions
    {
        /// <summary>
        /// 设置 单元格 数据格式
        /// 默认
        /// </summary>
        /// <param name="ICellStyle">单元格样式</param>
        /// <param name="dataFormat">HSSFDataFormat.GetBuiltinFormat("@")</param>
        /// <returns></returns>
        public static ICellStyle setDataFormat(this ICellStyle cellStyle, short dataFormat)
        {
            //为避免日期格式被Excel自动替换，所以设定 format 为 『@』 表示一率当成text來看
            cellStyle.DataFormat = dataFormat;

            return cellStyle;
        }

        ///// <summary>
        ///// ICellStyle
        ///// 设置 单元格 字体粗细-粗体
        ///// </summary>
        ///// <param name="ICellStyle">单元格样式</param>
        ///// <returns></returns>
        //public static ICellStyle setFontBold(this ICellStyle cellStyle)
        //{
        //    return cellStyle.setFont(FontBoldWeight.Bold);
        //}

        ///// <summary>
        ///// ICellStyle
        ///// 设置 单元格 字体粗细-正常值
        ///// </summary>
        ///// <param name="ICellStyle">单元格样式</param>
        ///// <returns></returns>
        //public static ICellStyle setFontNormal(this ICellStyle cellStyle)
        //{
        //    return cellStyle.setFont(FontBoldWeight.Normal);
        //}

        ///// <summary>
        ///// ICellStyle
        ///// 设置 单元格 字体粗细
        ///// </summary>
        ///// <param name="ICellStyle">单元格样式</param>
        ///// <param name="FontBoldWeight">字体粗细</param>
        ///// <returns></returns>
        //public static ICellStyle setFont(this ICellStyle cellStyle, FontBoldWeight fontBoldWeight)
        //{
        //    IFont font = workbook.CreateFont();

        //    font.Boldweight = (short)FontBoldWeight.Normal;
        //    cellStyle.SetFont(workbook.genFontNormal());

        //    return font;
        //}
    }
}