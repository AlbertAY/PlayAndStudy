using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExeclTool.Model
{
    /// <summary>
    /// execl的基础类
    /// </summary>
    public class ExeclBase
    {
        /// <summary>
        /// execl工作簿对象
        /// </summary>
        public IWorkbook BaseExcelWorkbook;
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="excelWorkbook"></param>
        public ExeclBase(IWorkbook excelWorkbook)
        {
            this.BaseExcelWorkbook = excelWorkbook;
        }

        /// <summary>
        /// 表头开始行索引,起始位置为0
        /// </summary>
        public int? TitleRowIndex { set; get; }
        /// <summary>
        /// 结果总数
        /// </summary>
        public int? ResultCount { set; get; }

        private HSSFFont _HSSFFont;
        /// <summary>
        ///execl的基础字体
        /// </summary>
        public HSSFFont BaseFont
        {
            set
            {
                _HSSFFont = value;
            }
            get
            {
                //如果没有设置基础字体则采用默认字体
                if (_HSSFFont == null)
                {
                    _HSSFFont = (HSSFFont)BaseExcelWorkbook.CreateFont();
                    _HSSFFont.FontHeightInPoints = 10;//字号 
                    _HSSFFont.FontName = "微软雅黑";
                    _HSSFFont.Color = NPOI.HSSF.Util.HSSFColor.Black.Index;//颜色 
                }
                return _HSSFFont;
            }
        }
        private ICellStyle _ICellStyle;
        /// <summary>
        /// 单元格基础样式
        /// </summary>
        public ICellStyle BaseCellStyle
        {
            set
            {
                _ICellStyle = value;
            }
            get
            {
                if (_ICellStyle == null)
                {
                    _ICellStyle = BaseExcelWorkbook.CreateCellStyle();
                    _ICellStyle.BorderBottom = BorderStyle.Thin;
                    _ICellStyle.BorderLeft = BorderStyle.Thin;
                    _ICellStyle.BorderRight = BorderStyle.Thin;
                    _ICellStyle.BorderTop = BorderStyle.Thin;
                    _ICellStyle.Alignment = HorizontalAlignment.Left;
                    _ICellStyle.SetFont(BaseFont);
                }
                return _ICellStyle;
            }
        }
        private ICellStyle _EditCellsStyle;
        /// <summary>
        /// 可编辑单元格样式
        /// </summary>
        public ICellStyle EditCellsStyle
        {
            set { _EditCellsStyle = value; }
            get
            {
                if (_EditCellsStyle != null)
                {
                    return _EditCellsStyle;
                }
                _EditCellsStyle = BaseCellStyle;
                _EditCellsStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightOrange.Index;
                _EditCellsStyle.FillPattern = FillPattern.SolidForeground;
                return _EditCellsStyle;
            }
        }
        private ICellStyle _NoEditCellsStyle;
        /// <summary>
        /// 不可编辑单元格样式
        /// </summary>
        public ICellStyle NoEditCellsStyle
        {
            set { _NoEditCellsStyle = value; }
            get
            {
                if (_NoEditCellsStyle != null)
                {
                    return _EditCellsStyle;
                }
                _NoEditCellsStyle = BaseCellStyle;
                //设置背景色
                _NoEditCellsStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
                _NoEditCellsStyle.FillPattern = FillPattern.SolidForeground;
                return _NoEditCellsStyle;
            }
        }


        private ICellStyle _NoEditTitleStyle;
        /// <summary>
        /// 非编辑列头部样式
        /// </summary>
        public ICellStyle NoEditTitleStyle
        {
            set { _NoEditTitleStyle = value; }
            get
            {
                if (_NoEditTitleStyle != null)
                {
                    return _NoEditTitleStyle;
                }
                _NoEditTitleStyle = BaseCellStyle;

                _NoEditTitleStyle.SetFont(BaseTitleFont);
                //设置背景色
                _NoEditTitleStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
                _NoEditTitleStyle.FillPattern = FillPattern.SolidForeground;
                return _NoEditTitleStyle;
            }         
        }

        private ICellStyle _EditTitleStyle;
        /// <summary>
        /// 编辑列头部样式
        /// </summary>
        public ICellStyle EditTitleStyle
        {
            set { _EditTitleStyle = value; }
            get
            {
                if (_EditTitleStyle != null)
                {
                    return EditTitleStyle;
                }
                _EditTitleStyle = BaseCellStyle;

                _EditTitleStyle.SetFont(BaseTitleFont);
                //设置背景色
                _EditTitleStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightOrange.Index;
                _EditTitleStyle.FillPattern = FillPattern.SolidForeground;
                return EditTitleStyle;
            }
        }


        private HSSFFont _BaseTitleFont;
        /// <summary>
        /// 表头基础样式
        /// </summary>
        public HSSFFont BaseTitleFont
        {
            set { _BaseTitleFont = value; }
            get
            {
                if (_BaseTitleFont != null)
                {
                    return _BaseTitleFont;
                }
                HSSFFont font = (HSSFFont)BaseExcelWorkbook.CreateFont();
                font.FontHeightInPoints = 10;//字号 
                font.FontName = "微软雅黑";
                font.Color = NPOI.HSSF.Util.HSSFColor.Black.Index;//颜色 
                font.Boldweight = 600;//HSSFFont.BOLDWEIGHT_BOLD;//加粗 
                return font;
            }
        }



        /// <summary>
        /// 基础背景色
        /// </summary>
        public short BaseBackGrandIndexed { get; set; }

        /// <summary>
        /// 工作簿是否加锁
        /// </summary>
        public bool WorkBookIsLock { get; set; }
    }
}
