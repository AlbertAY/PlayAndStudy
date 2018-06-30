using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExeclTool
{
    /// <summary>
    /// 冻结区域,详细参数说明参照NPOI的CreateFreezePane方法
    /// </summary>
    public class FreezePane
    {
        /// <summary>
        /// 冻结的列数
        /// </summary>
        public int ColSplit { set; get; }
        /// <summary>
        /// 冻结的行数
        /// </summary>
        public int RowSplit { set; get; }
        /// <summary>
        /// 右边区域可见的首列序号
        /// </summary>
        public int LeftmostColumn { set; get; }
        /// <summary>
        /// 下边区域可见的首行序号
        /// </summary>
        public int TopRow { set; get; }
    }
}
