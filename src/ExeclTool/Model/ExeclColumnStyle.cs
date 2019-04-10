using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExeclTool.Model
{
    //public class ExcelColumnStyle : ExcelBase
    //{
    //    /// <summary>
    //    /// 构造函数
    //    /// </summary>
    //    public ExcelColumnStyle(int? titleRowIndex, int? resultCount, IWorkbook excelWorkbook) : base(excelWorkbook)
    //    {
    //        this.ResultCount = resultCount;
    //        this.TitleRowIndex = titleRowIndex;
    //    }
    //    private int _width; //30 * 256
    //    private ICellStyle _CellStyle;
    //    private int? _StartBodyRow;
    //    private int? _EndBodyRow;

    //    /// <summary>
    //    /// 字段名称
    //    /// </summary>
    //    public string ColumnsName { set; get; }
    //    /// <summary>
    //    /// 是否锁定
    //    /// </summary>
    //    public bool IsLock { get; set; }
    //    /// <summary>
    //    /// 列宽度
    //    /// </summary>
    //    public int Width
    //    {
    //        get { return _width; }
    //        set
    //        {
    //            _width = value * 256;//单位是1/256个字符宽度，如果是30个字符长度就是 30*256。  
    //        }
    //    }
    //    /// <summary>
    //    /// 列计算公式
    //    /// </summary>
    //    public string SetCellFormula { set; get; }
    //    /// <summary>
    //    /// 列景色赋值
    //    /// </summary>
    //    public short BackGrandIndexed { get; set; }
    //    /// <summary>
    //    /// 是否隐藏列
    //    /// </summary>
    //    public bool IsHidden { set; get; }
    //    /// <summary>
    //    /// 表头的样式
    //    /// </summary>
    //    public ICellStyle TitleStyle { set; get; }
    //    /// <summary>
    //    /// 单元格式样式
    //    /// </summary>
    //    public ICellStyle CellStyle
    //    {
    //        set { _CellStyle = value; }
    //        get
    //        {
    //            //默认单元格样式
    //            return _CellStyle == null ? base.NoEditCellsStyle : _CellStyle;
    //        }
    //    }
    //    /// <summary>
    //    /// 单元格格式
    //    /// </summary>
    //    public string DataFormat { set; get; }
    //    /// <summary>
    //    /// 下拉框选择
    //    /// </summary>
    //    public string[] SelectValue { set; get; }
    //    /// <summary>
    //    /// 正文开始行
    //    /// </summary>
    //    public int StartBodyRow
    //    {
    //        set { _StartBodyRow = value; }
    //        get
    //        {
    //            if (_StartBodyRow.HasValue)
    //            {
    //                return _StartBodyRow.Value;
    //            }
    //            int startRow = TitleRowIndex.HasValue ? TitleRowIndex.Value + 1 : 0;
    //            return startRow;
    //        }
    //    }
    //    /// <summary>
    //    /// 正文结束行
    //    /// </summary>
    //    public int EndBodyRow
    //    {
    //        set { _EndBodyRow = value; }
    //        get
    //        {
    //            if (_EndBodyRow.HasValue)
    //            {
    //                return _EndBodyRow.Value;
    //            }
    //            return ResultCount.HasValue ? StartBodyRow + ResultCount.Value - 1 : StartBodyRow - 1;
    //        }
    //    }
    //    /// <summary>
    //    ///验证格式
    //    /// </summary>
    //    public DVConstraint DVConstraint { set; get; }
    //    /// <summary>
    //    /// 验证警示语
    //    /// </summary>
    //    public string DVConstraintWaringMsg { set; get; }
    //}
}
