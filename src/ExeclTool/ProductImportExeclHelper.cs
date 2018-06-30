
using ExeclTool.Model;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace ExeclTool
{
    /// <summary>
    /// 指标属性值
    /// </summary>
    public class ProductImportExeclHelper
    {
        /// <summary>
        /// 一个字节在excel中的宽度
        /// </summary>
        private static readonly int byteWidth = 650;
        /// <summary>
        /// 需要初始化设置的表头的列名（必须设置）
        /// </summary>
        public static Dictionary<string, string> ListColumnsName;

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="list">数据集合</param>
        /// <param name="execlWorkBookStyle">单元格样式集合</param>
        /// <param name="fileName">文件名称</param>
        public static void ExportExcel<T>(IList<T> list, WorkBookStyle execlWorkBookStyle, string fileName)
        {
            //当前请求上下文
            HttpResponse response = HttpContext.Current.Response;
            //生成名称
            fileName = string.Format("attachment;filename={0}({1}).xls", fileName, DateTime.Now.ToString("yyyyMMdd"));
            //生成文件流
            MemoryStream excelStream = ExportExcel(list, execlWorkBookStyle);
            response.Clear();
            //表头采用gb3212编码，解决在ie情况下的文件名乱码问题
            response.HeaderEncoding = System.Text.Encoding.GetEncoding("utf-8");//表头添加编码格式 
            response.AddHeader("content-disposition", fileName);
            response.Charset = "utf-8";
            //响应体采用gb2312编码
            response.ContentEncoding = System.Text.Encoding.GetEncoding("GB2312");
            response.ContentType = "application/ms-excel";
            //填充文件流
            response.BinaryWrite((excelStream).ToArray());
            excelStream.Close();
            //流释放
            excelStream.Dispose();
            response.End();
        }
        /// <summary>
        /// 生成execl文件流
        /// </summary>
        /// <param name="list">数据集合</param>
        /// <param name="execlWorkBookStyle">单元格样式集合</param>
        /// <returns></returns>
        public static MemoryStream ExportExcel<T>(IList<T> list, WorkBookStyle execlWorkBookStyle)
        {
            ListColumnsName = execlWorkBookStyle.ColumnsName;
            //判断列名是否存在
            if (ListColumnsName == null || ListColumnsName.Count == 0)
            {
                throw (new Exception("请对ListColumnsName设置要导出的列名！"));
            }
            MemoryStream excelStream = new MemoryStream();
            //表填充到execl
            InsertRow(list, execlWorkBookStyle);
            ExcelHelper.SaveExcelFile(execlWorkBookStyle.BaseExcelWorkbook, excelStream);
            return excelStream;
        }


        /// <summary>
        /// 插入数据行
        /// </summary>
        /// <param name="dtSource">数据表</param>
        /// <param name="execlWorkBookStyle">execl单元格样式</param>
        public static void InsertRow<T>(IList<T> dtSource, WorkBookStyle execlWorkBookStyle)
        {
            IWorkbook excelWorkbook = execlWorkBookStyle.BaseExcelWorkbook;
            int rowCount = execlWorkBookStyle.TitleRowIndex.HasValue ? execlWorkBookStyle.TitleRowIndex.Value : 0;
            //数据源导出数据集
            ISheet newsheet = execlWorkBookStyle.CreateWorkSheet(execlWorkBookStyle.SheetName);
            //警告区域
            ExcelHelper.CreateAreaBlock(execlWorkBookStyle);
            ICellStyle style = ExcelHelper.BorderCellStyle(excelWorkbook);
            //设置表头
            CreateHeader(newsheet, style, execlWorkBookStyle);
            foreach (T dr in dtSource)
            {
                rowCount++;
                //创建行
                IRow newRow = newsheet.CreateRow(rowCount);
                //行数据添加
                InsertCell(dr, newRow, execlWorkBookStyle, rowCount);
            }
            //设置工作簿是否加锁
            ExcelHelper.SetExcelSheetLock(execlWorkBookStyle.WorkBookIsLock, newsheet);
            //当前工作薄添加数字限制区域
            ExeclHelperExtend.SheetAddValidationArea(newsheet, execlWorkBookStyle, ListColumnsName);
            //设置选择下拉框
            ExeclHelperExtend.SetSelectValueArea(execlWorkBookStyle.ExeclColumnStyleList, ListColumnsName, newsheet);

            //冻结列
            if (execlWorkBookStyle.FreezePane != null)
            {
                FreezePane freezePane = execlWorkBookStyle.FreezePane;
                execlWorkBookStyle.WorkSheet.CreateFreezePane(freezePane.ColSplit, freezePane.RowSplit, freezePane.LeftmostColumn, freezePane.TopRow);
            }
        }
        /// <summary>
        /// 导出数据行
        /// </summary>
        /// <param name="drSource">来源数据行</param>
        /// <param name="currentExcelRow">当前数据行</param>
        /// <param name="execlWorkBookStyle">表头样式</param>
        /// <param name="rowCount">execl当前行</param>
        private static void InsertCell<T>(T drSource, IRow currentExcelRow, WorkBookStyle execlWorkBookStyle, int rowCount)
        {
            Type objType = typeof(T);
            #region 获取错误信息集合
            List<ErrorMessage> errorMessages = new List<ErrorMessage>();
            //得到错误信息集合对象
            PropertyInfo errorMessagesPro = objType.GetProperty("ErrorMessages");
            if (errorMessagesPro != null)
            {
                errorMessages = (List<ErrorMessage>)errorMessagesPro.GetValue(drSource, null);
            }
            #endregion

            #region 获取未定义的列信息集合
            List<ColumnInfo> otherColumns = new List<ColumnInfo>();
            //得到未定义的列集合对象
            PropertyInfo otherColumnsPro = objType.GetProperty("OtherColumns");
            if (otherColumnsPro != null)
            {
                otherColumns = (List<ColumnInfo>)otherColumnsPro.GetValue(drSource, null);
            }
            #endregion

            int cellIndex = 0;
            //根据表头开始循环
            foreach (var item in ListColumnsName)
            {
                //列名称
                string columnsName = item.Key.ToString();
                PropertyInfo columnPropertyInfo = objType.GetProperty(columnsName);
                string drValue = string.Empty;
                System.Type rowType;
                #region 未定义属性处理
                if (columnPropertyInfo == null)
                {
                    ColumnInfo columnInfo = otherColumns.FirstOrDefault(it => it.Title.Equals(item.Value));
                    if (columnInfo != null)
                    {
                        drValue = columnInfo.Value;
                    }
                    columnsName = item.Value;
                    rowType = typeof(string);
                }
                else
                {
                    drValue = columnPropertyInfo.GetValue(drSource, null)?.ToString().Trim() ?? string.Empty;
                    rowType = columnPropertyInfo.PropertyType;
                }
                #endregion
                //获取列的样式对象
                ExeclColumnStyle execlCellStyle = execlWorkBookStyle.ExeclColumnStyleList?.FirstOrDefault(it => it.ColumnsName == columnsName) ?? null;
                ICell newCell = null;
                //获取该列的错误信息
                string errMsg = string.Empty;
                ErrorMessage errorMessage = errorMessages.FirstOrDefault(it => (it.Name?.Equals(columnsName) ?? false)
                                                                            || (it.Title?.Equals(columnsName) ?? false));
                //获取错误信息
                errMsg = errorMessage != null ? errorMessage.ErrMsg : string.Empty;
                
                newCell = currentExcelRow.CreateCell(cellIndex);
                //单元格内容处理
                CellCreate(rowType, newCell, drValue, execlWorkBookStyle.BaseExcelWorkbook);
                //设置列样式                
                ExcelHelper.SetExeclCellStyle(newCell, execlCellStyle);

                //设置批注                
                ExcelHelper.SetComment(errMsg, currentExcelRow, newCell, execlWorkBookStyle);

                //设置计算表达式
                ExcelHelper.SetCellFormula(execlCellStyle, execlWorkBookStyle.WorkSheet, newCell, rowCount);
                cellIndex++;
            }
        }
        /// <summary>
        /// 创建单元格值
        /// </summary>
        /// <param name="rowType">需要填入单元格字段的值的类型</param>
        /// <param name="newCell">单元格对象</param>
        /// <param name="drValue">单元格需要填入的值</param>
        /// <param name="workbook">工作簿对象</param>
        private static void CellCreate(Type rowType, ICell newCell, string drValue, IWorkbook workbook)
        {
            switch (rowType.ToString())
            {

                case "System.String"://字符串类型
                    drValue = drValue.Replace("&", "&");
                    drValue = drValue.Replace(">", ">");
                    drValue = drValue.Replace("<", "<");
                    newCell.SetCellValue(drValue);
                    break;
                case "System.Guid"://GUID类型                        
                    newCell.SetCellValue(drValue);
                    break;
                case "System.DateTime"://日期类型
                    DateTime dateV;
                    DateTime.TryParse(drValue, out dateV);
                    newCell.SetCellValue(dateV);
                    //格式化显示
                    ICellStyle cellStyle = workbook.CreateCellStyle();
                    IDataFormat format = workbook.CreateDataFormat();
                    cellStyle.DataFormat = format.GetFormat("yyyy-mm-dd hh:mm:ss");
                    newCell.CellStyle = cellStyle;
                    break;
                case "System.Boolean"://布尔型
                    bool boolV = false;
                    bool.TryParse(drValue, out boolV);
                    newCell.SetCellValue(boolV);
                    break;
                case "System.Int16"://整型
                case "System.Int32":
                case "System.Int64":
                case "System.Byte":
                    int intV = 0;
                    int.TryParse(drValue, out intV);
                    //设置值
                    newCell.SetCellValue(intV.ToString());
                    break;
                case "System.Decimal"://浮点型
                case "System.Double":
                    double doubV = 0;
                    double.TryParse(drValue, out doubV);
                    newCell.SetCellValue(doubV);
                    break;
                case "System.DBNull"://空值处理
                    newCell.SetCellValue("");
                    break;
                default:
                    newCell.SetCellValue(drValue);
                    break;
            }
        }

        /// <summary>
        /// 创建表头
        /// </summary>
        /// <param name="excelSheet"></param>
        /// <param name="style"></param>
        /// <param name="execlWorkBookStyle"></param>
        internal static void CreateHeader(ISheet excelSheet, ICellStyle style, WorkBookStyle execlWorkBookStyle)
        {
            List<ExeclColumnStyle> execlCellStyleList = execlWorkBookStyle.ExeclColumnStyleList;
            int cellIndex = 0, rowIndex;
            //title所在行
            rowIndex = execlWorkBookStyle.TitleRowIndex.HasValue ? execlWorkBookStyle.TitleRowIndex.Value : 0;
            IRow newRow = excelSheet.CreateRow(rowIndex);
            //循环导出列
            foreach (var de in ListColumnsName)
            {
                ICell newCell = newRow.CreateCell(cellIndex);
                newCell.SetCellValue(de.Value.ToString());
                //按表头文字的宽度
                excelSheet.SetColumnWidth(cellIndex, de.Value.ToString().Length * byteWidth);
                ExeclColumnStyle execlCellStyle = execlCellStyleList.FirstOrDefault(it => it.ColumnsName == de.Key);
                if (execlCellStyle != null)
                {
                    newCell.CellStyle = execlCellStyle.TitleStyle;//用户自定义表头样式                    
                }
                else
                {
                    newCell.CellStyle = style;//赋默认值
                }
                if ((execlCellStyle?.Width ?? 0) > 0)
                {
                    excelSheet.SetColumnWidth(cellIndex, execlCellStyle.Width);
                }
                cellIndex++;
            }
        }
    }
}
