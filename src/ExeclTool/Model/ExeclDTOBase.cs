using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExeclTool
{
    /// <summary>
    /// execl导出基础对象
    /// </summary>
    public class ExeclDTOBase
    {
        /// <summary>
        /// 对象中没有属性的一些其他列
        /// </summary>
        public List<ColumnInfo> OtherColumns { set; get; }
        /// <summary>
        /// 错误信息
        /// </summary>
        public List<ErrorMessage> ErrorMessages { set; get; }

    }
    /// <summary>
    /// 列信息
    /// </summary>
    public class ColumnInfo
    {
        /// <summary>
        /// execl的表头
        /// </summary>
        public string Title { set; get; }
        /// <summary>
        /// 该表头对于的属性名称
        /// </summary>
        public string Name { set; get; }
        /// <summary>
        /// 该属性对于的值
        /// </summary>
        public string Value { set; get; }

        /// <summary>
        /// 导入时候的列的位置
        /// </summary>
        public int ColumnIndex { set; get; }

    }
    /// <summary>
    /// 错误信息
    /// </summary>
    public class ErrorMessage
    {
        /// <summary>
        /// execl的表头
        /// </summary>
        public string Title { set; get; }
        /// <summary>
        /// 该表头对于的属性名称
        /// </summary>
        public string Name { set; get; }

        /// <summary>
        /// 错误描述
        /// </summary>
        public string ErrMsg { set; get; }
    }

}
