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
    /// 工作簿的基本信息
    /// </summary>
    public class WorkBookStyle : ExeclBase
    {
        /// <summary>
        /// execl工作簿
        /// </summary>
        public static IWorkbook ExcelWorkbook;
        /// <summary>
        /// 0参构造函数
        /// </summary>
        public WorkBookStyle() : base(ExcelWorkbook)
        {
            ExcelWorkbook = new HSSFWorkbook();
            base.BaseExcelWorkbook = ExcelWorkbook;
            //初始化相关参数
            ExeclColumnStyleList = new List<ExeclColumnStyle>();
        }
        /// <summary>
        /// 初始化相关参数
        /// </summary>
        /// <param name="columnsName">execl表头key是字段名，value是execl显示的title</param>

        public WorkBookStyle(Dictionary<string, string> columnsName) : base(ExcelWorkbook)
        {
            ExcelWorkbook = new HSSFWorkbook();
            base.BaseExcelWorkbook = ExcelWorkbook;
            //初始化相关参数
            ExeclColumnStyleList = new List<ExeclColumnStyle>();
            _columnsName = columnsName;
        }

        /// <summary>
        /// 构造函数，初始化标题和导出数量总数
        /// </summary>
        /// <param name="titleRowIndex">表头开始行</param>
        /// <param name="resultCount">结果总数</param>
        /// <param name="columnsName">execl表头key是字段名，value是execl显示的title</param>
        public WorkBookStyle(int titleRowIndex, int resultCount, Dictionary<string, string> columnsName) : base(ExcelWorkbook)
        {
            ExcelWorkbook = new HSSFWorkbook();
            base.BaseExcelWorkbook = ExcelWorkbook;
            this.ResultCount = resultCount;
            this.TitleRowIndex = titleRowIndex;
            this.ExeclColumnStyleList = new List<ExeclColumnStyle>();
            ColumnsName = columnsName;
        }
        private Dictionary<string, string> _columnsName;
        /// <summary>
        /// execl工作簿表头列
        /// </summary>
        public Dictionary<string, string> ColumnsName
        {
            set { _columnsName = value; }
            get
            {
                return _columnsName;
            }
        }
        /// <summary>
        /// 冻结列区域
        /// </summary>
        public FreezePane FreezePane { set; get; }
        /// <summary>
        /// execl上的其他区域
        /// </summary>
        public List<AreaBlock> AreaBlockList { set; get; }

        /// <summary>
        /// execl的列对象格式
        /// </summary>
        public List<ExeclColumnStyle> ExeclColumnStyleList { set; get; }
        /// <summary>
        /// 当前execl的工作页
        /// </summary>
        public ISheet WorkSheet { set; get; }
        private string _SheetName;
        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName { set { _SheetName = value; } get { return string.IsNullOrEmpty(_SheetName) ? "sheet1" : _SheetName; } }
        /// <summary>
        /// 列对象创建公共方法
        /// </summary>
        /// <param name="columnName">列名称</param>
        /// <returns></returns>
        public ExeclColumnStyle ExeclColumnStyleCreate(string columnName)
        {
            ExeclColumnStyle execlColumnStyle = new ExeclColumnStyle(this.TitleRowIndex, this.ResultCount, BaseExcelWorkbook);
            execlColumnStyle.ColumnsName = columnName;
            return execlColumnStyle;
        }
        /// <summary>
        ///  创建当前的工作簿工作表
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public ISheet CreateWorkSheet(string sheetName)
        {
            ISheet newSheet = ExcelWorkbook.CreateSheet(sheetName);
            WorkSheet = newSheet;
            SheetName = sheetName;
            return newSheet;
        }

    }
}
