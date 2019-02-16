using System;
using System.Collections.Generic;
using System.Linq;

namespace BlueColor.NPOIExtensions
{
    /// <summary>
    /// 属性单元格配置
    /// </summary>
    public class PropertyCellConfig
    {
        /// <summary>
        /// 属性名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 列索引（不能重复）
        /// </summary>
        public int ColumnIndex { get; set; } = -1;

        /// <summary>
        /// 条件列（最多只能有一列）
        /// </summary>
        public bool ConditionColumn { get; set; }

        /// <summary>
        /// 允许合并(ConditionColumn存在时，必须至少有一列AllowMerge)
        /// </summary>
        public bool AllowMerge { get; set; }

        /// <summary>
        /// 被忽略了
        /// </summary>
        public bool IsIgnored { get; set; }

        /// <summary>
        /// 数据格式
        /// </summary>
        public string DataFormat { get; set; } = "@";
    }

    /// <summary>
    /// 属性单元格配置List
    /// </summary>
    public class PropertyCellConfigList : List<PropertyCellConfig>
    {
        private bool IsValidated = false;

        /// <summary>
        /// 检验错误
        /// </summary>
        public void ValidateError()
        {
            if (this.IsValidated)
            {
                // 只验证一次就好！
                return;
            }
            else
            {
                this.IsValidated = true;
            }

            if (this.Count == 0)
            {
                throw new Exception("PropertyCellConfigList.Count == 0，请先设置！");
            }

            var configsColumnIndexs = this.Where(p => p.ColumnIndex != -1).Select(p => p.ColumnIndex).ToList();
            if (configsColumnIndexs.Count > 0)
            {
                if (configsColumnIndexs.Count != configsColumnIndexs.Distinct().Count())
                {
                    throw new Exception("ColumnIndex存在重复项，请检查！");
                }
            }

            var configConditionColumns = this.Where(p => p.ConditionColumn);
            if (configConditionColumns.Count() > 1)
            {
                throw new Exception("ConditionColumn最多只能有一列，请检查！");
            }
            else if (configConditionColumns.Count() == 1)
            {
                if (!this.Any(p => p.AllowMerge))
                {
                    throw new Exception("ConditionColumn（合并条件列）存在时，至少有一列AllowMerge为True，请设置！");
                }
            }

            // IsIgnored属性Bug影响未检验。
        }
    }
}