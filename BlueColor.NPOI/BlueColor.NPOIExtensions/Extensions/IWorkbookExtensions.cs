using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace BlueColor.NPOIExtensions.Extensions
{
    /// <summary>
    /// IWorkbook 扩展
    /// </summary>
    public static class IWorkbookExtensions
    {
        /// <summary>
        /// 生成 标题单元格样式
        /// 默认
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <returns></returns>
        public static ICellStyle genHeaderCellStyle(this IWorkbook workbook)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();

            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;

            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;

            cellStyle.SetFont(workbook.genFontBold());

            return cellStyle;
        }

        /// <summary>
        /// 生成 内容单元格样式
        /// 默认
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <returns></returns>
        public static ICellStyle genContentCellStyle(this IWorkbook workbook)
        {
            ICellStyle cellStyle = workbook.CreateCellStyle();

            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;

            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;

            //为避免日期格式被Excel自动替换，所以设定 format 为 『@』 表示一率当成text來看
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");

            cellStyle.SetFont(workbook.genFontNormal());

            return cellStyle;
        }

        /// <summary>
        /// 生成 字体-粗体
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <returns></returns>
        public static IFont genFontBold(this IWorkbook workbook)
        {
            return workbook.genFontBoldWeight(FontBoldWeight.Bold);
        }

        /// <summary>
        /// 生成 字体-正常值
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <returns></returns>
        public static IFont genFontNormal(this IWorkbook workbook)
        {
            return workbook.genFontBoldWeight(FontBoldWeight.Normal);
        }

        /// <summary>
        /// 生成 字体-指定粗细
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <param name="fontBoldWeight">字体粗体重量</param>
        /// <returns></returns>
        public static IFont genFontBoldWeight(this IWorkbook workbook, FontBoldWeight fontBoldWeight)
        {
            IFont font = workbook.CreateFont();
            font.Boldweight = (short)fontBoldWeight;

            return font;
        }
    }
}