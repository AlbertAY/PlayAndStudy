using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExeclTool.Model
{
    /// <summary>
    /// 列样式对象
    /// </summary>
    public class ExeclColumnStyle : ExeclBase
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public ExeclColumnStyle(int? titleRowIndex, int? resultCount, IWorkbook excelWorkbook) : base(excelWorkbook)
        {
            this.ResultCount = resultCount;
            this.TitleRowIndex = titleRowIndex;
        }
        private int _width; //30 * 256
        /// <summary>
        /// 字段名称
        /// </summary>
        public string ColumnsName { set; get; }
        /// <summary>
        /// 默认不加锁
        /// </summary>
        private bool _IsLock = false;


        /// <summary>
        /// 是否锁定
        /// </summary>
        public bool IsLock { get { return _IsLock; } set { _IsLock = value; } }
        /// <summary>
        /// 列宽度
        /// </summary>

        public int Width
        {
            get { return _width; }
            set
            {
                _width = value * 256;//单位是1/256个字符宽度，如果是30个字符长度就是 30*256。  
            }
        }
        /// <summary>
        /// 列计算公式
        /// </summary>
        public string SetCellFormula { set; get; }

        /// <summary>
        /// 列景色赋值
        /// </summary>
        public short BackGrandIndexed { get; set; }

        private bool _IsHidden = false;

        /// <summary>
        /// 是否隐藏列
        /// </summary>
        public bool IsHidden { set { _IsHidden = value; } get { return _IsHidden; } }
        private ICellStyle _TitleStyle;
        /// <summary>
        /// 表头的样式
        /// </summary>
        public ICellStyle TitleStyle
        {
            set { _TitleStyle = value; }
            get
            {
                //默认表头样式
                return _TitleStyle == null ? base.EditTitleStyle : _TitleStyle;
            }
        }


        private ICellStyle _CellStyle;
        /// <summary>
        /// 单元格式样式
        /// </summary>

        public ICellStyle CellStyle
        {
            set { _CellStyle = value; }
            get
            {
                //默认单元格样式
                return _CellStyle == null ? base.NoEditCellsStyle : _CellStyle;
            }
        }

        /// <summary>
        /// 单元格格式
        /// </summary>
        public string DataFormat { set; get; }
        /// <summary>
        /// 下拉框选择
        /// </summary>
        public string[] SelectValue { set; get; }
        /// <summary>
        /// 区域验证类型
        /// </summary>
        public string ValidationType { set; get; }

        private int? _StartBodyRow;
        /// <summary>
        /// 正文开始行
        /// </summary>
        public int StartBodyRow
        {
            set { _StartBodyRow = value; }
            get
            {
                if (_StartBodyRow.HasValue)
                {
                    return _StartBodyRow.Value;
                }
                int startRow = TitleRowIndex.HasValue ? TitleRowIndex.Value + 1 : 0;
                return startRow;
            }
        }

        private int? _EndBodyRow;
        /// <summary>
        /// 正文结束行
        /// </summary>
        public int EndBodyRow
        {
            set { _EndBodyRow = value; }
            get
            {
                if (_EndBodyRow.HasValue)
                {
                    return _EndBodyRow.Value;
                }
                return ResultCount.HasValue ? StartBodyRow + ResultCount.Value - 1 : StartBodyRow - 1;
            }
        }
    }
}
