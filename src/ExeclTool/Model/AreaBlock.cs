using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExeclTool.Model
{
    /// <summary>
    /// Execl区块
    /// </summary>
    public class AreaBlock
    {
        /// <summary>
        /// 开始行
        /// </summary>
        public int Rowstart { set; get; }
        /// <summary>
        /// 结束行
        /// </summary>
        public int Rowend { get; set; }
        /// <summary>
        /// 开始列
        /// </summary>
        public int Colstart { get; set; }
        /// <summary>
        /// 结束列
        /// </summary>
        public int Colend { set; get; }
        /// <summary>
        /// 高度
        /// </summary>
        public short? Height { set; get; }
        /// <summary>
        /// 内容
        /// </summary>
        public string Content { set; get; }
        /// <summary>
        /// 当前区域的样式
        /// </summary>
        public ICellStyle CellStyle { set; get; }
    }
}
