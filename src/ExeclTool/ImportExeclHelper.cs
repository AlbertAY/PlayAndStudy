using ExeclTool.Model;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace ExeclTool
{
    class ImportExeclHelper
    {
        /// <summary>
        /// 一个字节在excel中的宽度
        /// </summary>
        private static readonly int byteWidth = 650;

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
            //判断列名是否存在
            if (execlWorkBookStyle.ColumnsName == null || execlWorkBookStyle.ColumnsName.Count == 0)
            {
                //获取特性的表头
                Dictionary<string, string> title = DataHelper.GetExeclTitleByTitleAttribute<T>();
                if (title.Count <= 0)
                {
                    throw (new Exception("请对ListColumnsName设置要导出的列名！"));
                }
                execlWorkBookStyle.ColumnsName = title;
            }
            MemoryStream excelStream = new MemoryStream();
            //表填充到execl
            InsertRow(list, execlWorkBookStyle);
            ExcelHelper.SaveExcelFile(execlWorkBookStyle.BaseExcelWorkbook, excelStream);
            return excelStream;
        }

        /// <summary>
        /// 生成execl文件流
        /// </summary>
        /// <param name="list">数据集合</param>
        /// <param name="execlWorkBookStyle">单元格样式集合</param>
        /// <returns></returns>
        public static IWorkbook GetWorkbook<T>(IList<T> list, WorkBookStyle execlWorkBookStyle)
        {
            //判断列名是否存在
            if (execlWorkBookStyle.ColumnsName == null || execlWorkBookStyle.ColumnsName.Count == 0)
            {
                throw (new Exception("请对ListColumnsName设置要导出的列名！"));
            }
            MemoryStream excelStream = new MemoryStream();
            //表填充到execl
            InsertRow(list, execlWorkBookStyle);
            return execlWorkBookStyle.BaseExcelWorkbook;
        }



        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="execlWorkBookStyle">单元格样式集合</param>
        /// <param name="fileName">文件名称</param>
        public static void ExportExcel(MemoryStream stream, WorkBookStyle execlWorkBookStyle, string fileName)
        {
            //当前请求上下文
            HttpResponse response = HttpContext.Current.Response;
            //生成名称
            fileName = string.Format("attachment;filename={0}({1}).xls", fileName, DateTime.Now.ToString("yyyyMMdd"));
            //生成文件流
            MemoryStream excelStream = stream;
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
        /// 插入数据行
        /// </summary>
        /// <param name="dtoList">数据对象集合</param>
        /// <param name="execlWorkBookStyle">execl单元格样式</param>
        public static void InsertRow<T>(IList<T> dtoList, WorkBookStyle execlWorkBookStyle)
        {
            IWorkbook excelWorkbook = execlWorkBookStyle.BaseExcelWorkbook;
            int titleRowIndex = execlWorkBookStyle.TitleRowIndex.HasValue ? execlWorkBookStyle.TitleRowIndex.Value : 0;
            ISheet newsheet = execlWorkBookStyle.WorkSheet;
            //数据源导出数据集
            if (newsheet == null)
            {
                newsheet = execlWorkBookStyle.CreateWorkSheet(execlWorkBookStyle.SheetName);
            }
            //警告区域
            ExcelHelper.CreateAreaBlock(execlWorkBookStyle);
            ICellStyle style = ExcelHelper.BorderCellStyle(excelWorkbook);
            //设置表头
            CreateHeader(newsheet, style, execlWorkBookStyle);
            //行游标
            int rowCount = titleRowIndex + 1;
            foreach (T entity in dtoList)
            {
                int createRowIndex;

                //获取该对象的
                int? indexValue = GetEntityRowIndexValue(entity);
                if (indexValue.HasValue == false)
                {
                    createRowIndex = rowCount;
                    rowCount++;
                }
                else
                {
                    //源execl该数据所在行
                    createRowIndex = indexValue.Value + titleRowIndex + 1;
                }
                IRow newRow = newsheet.GetRow(createRowIndex);
                if (newRow == null)
                {
                    //创建行
                    newRow = newsheet.CreateRow(createRowIndex);
                }
                //行数据添加
                InsertCell(entity, newRow, execlWorkBookStyle, rowCount);
            }
            //设置工作簿是否加锁
            ExcelHelper.SetExcelSheetLock(execlWorkBookStyle.WorkBookIsLock, newsheet);
            //当前工作薄添加数字限制区域
            ExeclHelperExtensions.SheetAddValidationArea(newsheet, execlWorkBookStyle);
            //设置选择下拉框
            ExeclHelperExtensions.SetSelectValueArea(execlWorkBookStyle.ExeclColumnStyleList, execlWorkBookStyle.ColumnsName, newsheet);

            //冻结列
            if (execlWorkBookStyle.FreezePane != null)
            {
                FreezePane freezePane = execlWorkBookStyle.FreezePane;
                execlWorkBookStyle.WorkSheet.CreateFreezePane(freezePane.ColSplit, freezePane.RowSplit, freezePane.LeftmostColumn, freezePane.TopRow);
            }
        }

        /// <summary>
        /// 获取的实体的行标记
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="entity"></param>
        /// <returns></returns>
        private static int? GetEntityRowIndexValue<T>(T entity)
        {
            //为空判断
            if (entity == null)
            {
                return null;
            }
            //获取实体的指定属性
            PropertyInfo rowIndex = entity.GetType().GetProperty("RowIndex");
            if (rowIndex == null)
            {
                return null;
            }
            //获取属性的值
            string rowIndexValue = rowIndex?.GetValue(entity, null)?.ToString();
            if (string.IsNullOrEmpty(rowIndexValue))
            {
                return null;
            }
            int convertValue;
            //值类型强转
            if (int.TryParse(rowIndexValue, out convertValue))
            {
                return convertValue;
            }

            //默认返回null
            return null;
        }

        /// <summary>
        /// 插入数据行
        /// </summary>
        /// <param name="entity">来源数据行</param>
        /// <param name="currentExcelRow">当前数据行</param>
        /// <param name="execlWorkBookStyle">表头样式</param>
        /// <param name="rowCount">execl当前行</param>
        private static void InsertCell<T>(T entity, IRow currentExcelRow, WorkBookStyle execlWorkBookStyle, int rowCount)
        {
            Type objType = typeof(T);
            #region 获取错误信息集合
            List<ErrorMessage> errorMessages = new List<ErrorMessage>();
            //得到错误信息集合对象
            PropertyInfo errorMessagesPro = objType.GetProperty("ErrorMessages");
            if (errorMessagesPro != null)
            {
                errorMessages = (List<ErrorMessage>)errorMessagesPro.GetValue(entity, null);
            }
            #endregion

            #region 获取未定义的列信息集合
            List<ColumnInfo> otherColumns = new List<ColumnInfo>();
            //得到未定义的列集合对象
            PropertyInfo otherColumnsPro = objType.GetProperty("OtherColumns");
            if (otherColumnsPro != null)
            {
                otherColumns = (List<ColumnInfo>)otherColumnsPro.GetValue(entity, null);
            }
            #endregion

            int cellIndex = 0;
            //根据表头开始循环
            foreach (var item in execlWorkBookStyle.ColumnsName)
            {
                //列名称
                string columnsName = item.Key.ToString();
                PropertyInfo columnPropertyInfo = objType.GetProperty(columnsName);
                string propertyValue = string.Empty;
                System.Type rowType;
                #region 未定义属性处理
                if (columnPropertyInfo == null)
                {
                    ColumnInfo columnInfo = otherColumns.FirstOrDefault(it => it.Title.Equals(item.Value));
                    if (columnInfo != null)
                    {
                        propertyValue = columnInfo.Value;
                    }
                    columnsName = item.Value;
                    rowType = typeof(string);
                }
                else
                {
                    propertyValue = columnPropertyInfo.GetValue(entity, null)?.ToString().Trim() ?? string.Empty;
                    rowType = columnPropertyInfo.PropertyType;
                }
                #endregion
                //获取列的样式对象
                ExcelColumnStyle execlCellStyle = execlWorkBookStyle.ExeclColumnStyleList?.FirstOrDefault(it => it.ColumnsName == columnsName) ?? null;
                //获取该列的错误信息
                string errMsg = string.Empty;
                List<ErrorMessage> errorMessageList = errorMessages?.Where(it => (it.Name?.Equals(columnsName) ?? false)
                                                                            || (it.Title?.Equals(columnsName) ?? false))?.ToList();

                ICell newCell = currentExcelRow.GetCell(cellIndex);
                if (newCell == null)
                {
                    newCell = currentExcelRow.CreateCell(cellIndex);
                }
                else
                {
                    newCell.CellComment = null;
                }

                //单元格内容处理
                CellCreate(rowType, newCell, propertyValue, execlWorkBookStyle.BaseExcelWorkbook);
                //设置列样式                
                ExcelHelper.SetExeclCellStyle(newCell, execlCellStyle);
                //设置隐藏样式
                ExcelHelper.SetColumnStyle(execlWorkBookStyle.WorkSheet, cellIndex, execlCellStyle);
                //设置批注                
                SetComment(errorMessageList, currentExcelRow, newCell, execlWorkBookStyle);

                //设置计算表达式
                ExcelHelper.SetCellFormula(execlCellStyle, execlWorkBookStyle.WorkSheet, newCell, rowCount);
                cellIndex++;
            }
        }

        /// <summary>
        /// 设置批注
        /// </summary>
        /// <param name="errorMessageList"></param>
        /// <param name="row"></param>
        /// <param name="cell"></param>
        /// <param name="execlWorkBookStyle"></param>
        public static void SetComment(List<ErrorMessage> errorMessageList, IRow row, ICell cell, WorkBookStyle execlWorkBookStyle)
        {
            if (errorMessageList == null || errorMessageList.Count <= 0)
            {
                return;
            }
            //获取错误信息
            Array errmsgs = errorMessageList.Select(it => it.ErrMsg).ToArray();
            //多条错误信息换行处理
            string strMsg = String.Join(";\n", (string[])errmsgs);
            //设置批注内容
            ExcelHelper.SetComment(strMsg, row, cell, execlWorkBookStyle);

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
                case "System.Collections.Generic.List`1[Mysoft.Clgyl.Utility.Model.Excel.RowFile]":
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
        /// <param name="excelSheet">工作表</param>
        /// <param name="style">单元格样式</param>
        /// <param name="execlWorkBookStyle"></param>
        internal static void CreateHeader(ISheet excelSheet, ICellStyle style, WorkBookStyle execlWorkBookStyle)
        {
            List<ExeclColumnStyle> execlCellStyleList = execlWorkBookStyle.ExeclColumnStyleList;
            int cellIndex = 0, rowIndex;
            //title所在行
            rowIndex = execlWorkBookStyle.TitleRowIndex.HasValue ? execlWorkBookStyle.TitleRowIndex.Value : 0;
            IRow sheetRow = excelSheet.GetRow(rowIndex);
            if (sheetRow == null)
            {
                sheetRow = excelSheet.CreateRow(rowIndex);
            }
            //循环导出列
            foreach (var de in execlWorkBookStyle.ColumnsName)
            {
                ICell newCell = sheetRow.GetCell(cellIndex);
                if (newCell == null)
                {
                    newCell = sheetRow.CreateCell(cellIndex);
                    //赋默认值
                    newCell.CellStyle = style;
                    //按表头文字的宽度
                    excelSheet.SetColumnWidth(cellIndex, de.Value.ToString().Length * byteWidth);
                    //获取头部样式
                    ExcelColumnStyle execlCellStyle = execlCellStyleList.FirstOrDefault(it => it.ColumnsName == de.Key);
                    if (execlCellStyle != null && execlCellStyle.TitleStyle != null)
                    {
                        newCell.CellStyle = execlCellStyle.TitleStyle;//用户自定义表头样式                    
                    }
                    //设置头部宽度
                    if ((execlCellStyle?.Width ?? 0) > 0)
                    {
                        excelSheet.SetColumnWidth(cellIndex, execlCellStyle.Width);
                    }
                }
                newCell.SetCellValue(de.Value.ToString());

                cellIndex++;
            }
        }





    }
}

