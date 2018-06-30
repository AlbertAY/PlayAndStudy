using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations.Resources;

namespace ExeclTool
{
    /// <summary>
    /// execl列名称对象
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class ExeclTitleAttribute : ValidationAttribute
    {
        /// <summary>
        /// 构造函数，
        /// </summary>
        /// <param name="languageKey">表头标题</param>
        /// <param name="columnIndex">所在列排序</param>
        public ExeclTitleAttribute(string language, int columnIndex)
        {
            TitleName = language;
            ColumnIndex = columnIndex;
        }
        /// <summary>
        ///构造函数
        /// </summary>
        /// <param name="languageKey">多语言的key</param>
        public ExeclTitleAttribute(string language)
        {
            TitleName = language;
        }
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="columnIndex">列顺序</param>
        public ExeclTitleAttribute(int columnIndex)
        {
            ColumnIndex = columnIndex;
        }

        /// <summary>
        /// 列顺序
        /// </summary>
        public int? ColumnIndex { get; set; }

        /// <summary>
        /// 列名称
        /// </summary>
        public string TitleName { get; set; }
        /// <summary>
        /// 属性名称
        /// </summary>
        public string PropName { set; get; }
    }
}
