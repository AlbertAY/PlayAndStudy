//using ExeclTool.Model;
//using NPOI.HSSF.UserModel;
//using NPOI.HSSF.Util;
//using NPOI.SS.UserModel;
//using NPOI.SS.Util;
//using NPOI.XSSF.UserModel;
//using System;
//using System.Collections;
//using System.Collections.Generic;
//using System.Data;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Web;
//using System.Xml.Serialization;

//namespace ExeclTool
//{
//    /// <summary>
//    /// ExcelHelper帮助类
//    /// </summary>
//    public static class ExcelHelper
//    {
//        /// <summary>
//        /// 一个字节在excel中的宽度
//        /// </summary>
//        private static readonly int byteWidth = 650;

//        /// <summary>
//        /// execl的表头内容
//        /// </summary>
//        public static List<string> ExeclTitle { set; get; }
//        /// <summary>
//        ///     获取导入的workbook
//        /// </summary>
//        /// <param name="fs">文件流</param>
//        /// <param name="fileName">文件名</param>
//        /// <returns>导入的workbook</returns>
//        public static IWorkbook GetImportWorkBook(Stream fs, string fileName)
//        {
//            //工作簿
//            IWorkbook workbook = null;
//            if (fileName.IndexOf(".xlsx", StringComparison.Ordinal) > 0) // 2007版本
//            {
//                workbook = new XSSFWorkbook(fs);
//            }
//            else if (fileName.IndexOf(".xls", StringComparison.Ordinal) > 0) // 2003版本
//            {
//                workbook = new HSSFWorkbook(fs);
//            }
//            return workbook;
//        }

  

    

//        /// <summary>
//        ///     获取导入的sheet
//        /// </summary>
//        /// <param name="sheetName"></param>
//        /// <param name="workbook"></param>
//        /// <returns></returns>
//        public static ISheet GetImportSheet(IWorkbook workbook, string sheetName = null)
//        {
//            ISheet sheet;
//            if (sheetName != null)
//            {
//                //获取sheet
//                sheet = workbook.GetSheet(sheetName);
//            }
//            else
//            {
//                //获取第一个sheet
//                sheet = workbook.GetSheetAt(0);
//            }
//            return sheet;
//        }

        

//        /// <summary>
//        /// 根据列名获取列索引
//        /// </summary>
//        /// <param name="dataRow">行</param>
//        /// <param name="headName">列名</param>
//        /// <returns>列索引</returns>
//        public static int GetCellNum(IRow dataRow, string headName)
//        {
//            var cells = dataRow.Sheet.GetRow(1).Cells;
//            var columnIndex = 0;
//            foreach (var cell in cells)
//                if (cell.StringCellValue.Equals(headName))
//                {
//                    columnIndex = cell.ColumnIndex;
//                    break;
//                }

//            return columnIndex;
//        }

//        /// <summary>
//        /// 获取cell中的内容
//        /// </summary>
//        /// <param name="cellData"></param>
//        /// <returns></returns>
//        public static string GetCellContext(this ICell cellData)
//        {
//            switch (cellData.CellType)
//            {
//                case CellType.Numeric:
//                    return cellData.NumericCellValue.ToString();
//                case CellType.String:
//                    return cellData.StringCellValue;
//                case CellType.Boolean:
//                    return cellData.BooleanCellValue.ToString();
//                //case CellType.Formula: todo 公式处理
//                //     cellData.
//                default: return cellData.StringCellValue;
//            }
//        }

//        /// <summary>
//        /// 删除开始无效的列
//        /// </summary>
//        /// <param name="row"></param>
//        public static void RemoveHeadEmptyCell(this IRow row)
//        {
//            foreach (var cell in row.Cells)
//            {
//                if (string.IsNullOrEmpty(cell.GetCellContext().Trim()))
//                {
//                    row.RemoveCell(cell);
//                }
//                else
//                {
//                    break;
//                }
//            }
//        }

//        /// <summary>
//        /// 删除无效的列
//        /// </summary>
//        /// <param name="row"></param>
//        public static void RemoveEmptyCell(this IRow row)
//        {
//            foreach (var cell in row.Cells)
//            {
//                if (string.IsNullOrEmpty(cell.GetCellContext().Trim()))
//                {
//                    row.RemoveCell(cell);
//                }
//            }
//        }
//        /// <summary>
//        /// 根据文件信息获取文件的table
//        /// </summary>
//        /// <param name="fileInfoDto">文件基本信息</param>
//        /// <param name="sheetIndex">sheet页</param>
//        /// <param name="headerRowIndex">表头行</param>
//        /// <returns></returns>
//        public static Tuple<DataTable, IWorkbook, List<ISheet>> ImportExcel(string path, int sheetIndex, int headerRowIndex)
//        {
//            //文件判断
//            if (path == null )
//            {
//                //没有文件抛出异常
//                throw new Exception("");
//            }
//            //返回datatable对象
//            using (FileStream fs = File.OpenRead(path))
//            {
//                return ImportExcel(fs, sheetIndex, headerRowIndex);
//            }
            
            
//        }
//        /// <summary>
//        /// 根据文件信息获取文件的对象
//        /// </summary>
//        /// <typeparam name="T"></typeparam>
//        /// <param name="fileInfoDto"></param>
//        /// <param name="sheetIndex"></param>
//        /// <param name="headerRowIndex"></param>
//        /// <returns></returns>
//        public static Tuple<IList<T>, IWorkbook, List<ISheet>> GetExeclDTO<T>(Stream flie, int sheetIndex, int headerRowIndex)
//        {
//            //得到Execl中的数据表
//            Tuple<DataTable, IWorkbook,List<ISheet>> result= ExcelHelper.ImportExcel(flie, sheetIndex, headerRowIndex);
//            //将DataTable转成业务对象
//            IList<T> list =  DataHelper.ModelConver<T>(result.Item1);
//            return Tuple.Create(list, result.Item2, result.Item3);
//        }


           
//        public static Tuple<DataTable, IWorkbook, List<ISheet>> ImportExcel(Stream excelFileStream, int sheetIndex, int headerRowIndex)
//        {
//            List<string> title = new List<string>();
//            try
//            {
//                IWorkbook workbook;
//                try
//                {
//                    workbook = new XSSFWorkbook(excelFileStream);//2007版本
//                }
//                catch
//                {
//                    workbook = new HSSFWorkbook(excelFileStream);//2003版本
//                }

//                ISheet sheet = workbook.GetSheetAt(sheetIndex);
//                DataTable table = new DataTable();
//                //获取头部信息
//                IRow headerRow = sheet.GetRow(headerRowIndex);
//                int cellCount = headerRow.LastCellNum;
//                //获取头部信息
//                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
//                {
//                    string value = headerRow.GetCell(i).StringCellValue;
//                    DataColumn column = new DataColumn(value);
//                    title.Add(value);
//                    table.Columns.Add(column);
//                }
//                for (int i = (headerRowIndex + 1); i <= sheet.LastRowNum; i++)
//                {
//                    IRow row = sheet.GetRow(i);
//                    if (row == null)
//                        continue;

//                    DataRow dataRow = table.NewRow();
//                    for (int j = row.FirstCellNum; j < cellCount; j++)
//                    {
//                        dataRow[j] = row.GetCell(j) == null ? "" : row.GetCell(j).ToString().Trim();
//                    }

//                    //判断正行数据是否有效
//                    bool key = false;
//                    for (int ii = 0; ii < cellCount; ii++)
//                    {
//                        if (string.IsNullOrEmpty(dataRow[ii].ToString()) == false)
//                        {
//                            key = true;
//                            break;
//                        }
//                    }
//                    if (key)
//                    {
//                        table.Rows.Add(dataRow);
//                    }
//                }
//                List<ISheet> listSheet = new List<ISheet>();
//                listSheet.Add(sheet);
//                return Tuple.Create<DataTable, IWorkbook, List<ISheet>>(table, workbook, listSheet); 
//            }
//            catch (Exception e)
//            {
//                throw new Exception($"读取execl失败,错误信息：{e.Message}");
//            }
//            finally
//            {
//                excelFileStream.Close();
//            }
//        }

//        /// <summary>
//        /// 需要初始化设置的表头的列名（必须设置）
//        /// </summary>
//        public static Dictionary<string, string> ListColumnsName;


//        #region 导出

//        /// <summary>
//        /// 导出Excel
//        /// </summary>
//        /// <param name="list">数据集合</param>
//        /// <param name="execlWorkBookStyle">单元格样式集合</param>
//        /// <param name="fileName">文件名称</param>
//        public static void ExportExcel(IList list, WorkBookStyle execlWorkBookStyle, string fileName)
//        {
//            //当前请求上下文
//            HttpResponse response = HttpContext.Current.Response;
//            //生成名称
//            fileName = string.Format("attachment;filename={0}({1}).xls", fileName, DateTime.Now.ToString("yyyyMMdd"));
//            //生成文件流
//            MemoryStream excelStream = ExportExcel(list, execlWorkBookStyle);
//            response.Clear();
//            //表头采用gb3212编码，解决在ie情况下的文件名乱码问题
//            response.HeaderEncoding = System.Text.Encoding.GetEncoding("GB2312");//表头添加编码格式 
//            response.AddHeader("content-disposition", fileName);
//            response.Charset = "utf-8";
//            //响应体采用gb2312编码
//            response.ContentEncoding = System.Text.Encoding.GetEncoding("GB2312");
//            response.ContentType = "application/ms-excel";
//            //填充文件流
//            response.BinaryWrite((excelStream).ToArray());
//            excelStream.Close();
//            //流释放
//            excelStream.Dispose();
//            response.End();
//        }
//        /// <summary>
//        /// 生成execl文件流
//        /// </summary>
//        /// <param name="list">数据集合</param>
//        /// <param name="execlWorkBookStyle">单元格样式集合</param>
//        /// <returns></returns>
//        public static MemoryStream ExportExcel(IList list, WorkBookStyle execlWorkBookStyle)
//        {
//            ListColumnsName = execlWorkBookStyle.ColumnsName;
//            //判断列名是否存在
//            if (ListColumnsName == null || ListColumnsName.Count == 0)
//            {
//                throw (new Exception("请对ListColumnsName设置要导出的列名！"));
//            }
//            MemoryStream excelStream = new MemoryStream();
//            //获取table
//            DataTable dtSource = DataHelper.ToDataTable(list);
//            //表填充到execl
//            InsertRow(dtSource, execlWorkBookStyle);
//            SaveExcelFile(execlWorkBookStyle.BaseExcelWorkbook, excelStream);
//            return excelStream;
//        }


//        #endregion

//        #region 私有方法

//        /// <summary>
//        /// CellBorder设置（表头）
//        /// </summary>
//        /// <param name="excelWorkbook">工作簿</param>
//        /// <returns></returns>
//        internal static ICellStyle BorderCellStyle(IWorkbook excelWorkbook)
//        {
//            ICellStyle style = excelWorkbook.CreateCellStyle();
//            //单元格样式线
//            style.BorderBottom = BorderStyle.Thin;
//            style.BorderLeft = BorderStyle.Thin;
//            style.BorderRight = BorderStyle.Thin;
//            style.BorderTop = BorderStyle.Thin;
//            //居中
//            style.Alignment = HorizontalAlignment.Center;
//            //背景色
//            style.FillBackgroundColor = HSSFColor.White.Index;
//            return style;
//        }

//        /// <summary>
//        /// 保存Excel文件
//        /// </summary>
//        /// <param name="excelWorkBook">工作簿</param>
//        /// <param name="excelStream">文件流</param>
//        internal static void SaveExcelFile(IWorkbook excelWorkBook, Stream excelStream)
//        {
//            excelWorkBook.Write(excelStream);
//        }
//        /// <summary>
//        /// 创建Excel文件
//        /// </summary>
//        /// <returns></returns>
//        public static IWorkbook CreateExcelFile()
//        {
//            IWorkbook hssfworkbook = new HSSFWorkbook();
//            return hssfworkbook;
//        }
//        /// <summary>
//        /// 创建excel表头
//        /// </summary>
//        /// <param name="excelSheet">工作簿</param>
//        /// <param name="style">表头样式</param>
//        internal static void CreateHeader(ISheet excelSheet, ICellStyle style)
//        {
//            int cellIndex = 0;
//            IRow newRow = excelSheet.CreateRow(0);
//            //循环导出列
//            foreach (var de in ListColumnsName)
//            {
//                //创建单元格
//                ICell newCell = newRow.CreateCell(cellIndex);
//                newCell.SetCellValue(de.Value.ToString());
//                //按表头文字的宽度
//                excelSheet.SetColumnWidth(cellIndex, de.Value.ToString().Length * byteWidth);
//                newCell.CellStyle = style;
//                cellIndex++;
//            }
//        }
//        /// <summary>
//        /// 创建表头
//        /// </summary>
//        /// <param name="excelSheet"></param>
//        /// <param name="style"></param>
//        /// <param name="execlWorkBookStyle"></param>
//        internal static void CreateHeader(ISheet excelSheet, ICellStyle style, WorkBookStyle execlWorkBookStyle)
//        {
//            List<ExeclColumnStyle> execlCellStyleList = execlWorkBookStyle.ExeclColumnStyleList;
//            int cellIndex = 0, rowIndex;
//            //title所在行
//            rowIndex = execlWorkBookStyle.TitleRowIndex.HasValue ? execlWorkBookStyle.TitleRowIndex.Value : 0;
//            IRow newRow = excelSheet.CreateRow(rowIndex);
//            //循环导出列
//            foreach (var de in ListColumnsName)
//            {
//                ICell newCell = newRow.CreateCell(cellIndex);
//                newCell.SetCellValue(de.Value.ToString());
//                //按表头文字的宽度
//                excelSheet.SetColumnWidth(cellIndex, de.Value.ToString().Length * byteWidth);
//                ExeclColumnStyle execlCellStyle = execlCellStyleList.FirstOrDefault(it => it.ColumnsName == de.Key);
//                if (execlCellStyle != null)
//                {
//                    newCell.CellStyle = execlCellStyle.TitleStyle;//用户自定义表头样式
//                }
//                else
//                {
//                    newCell.CellStyle = style;//赋默认值
//                }
//                if ((execlCellStyle?.Width ?? 0) > 0)
//                {
//                    excelSheet.SetColumnWidth(cellIndex, execlCellStyle.Width);
//                }
//                cellIndex++;
//            }
//        }
//        /// <summary>
//        /// 插入数据行
//        /// </summary>
//        /// <param name="dtSource">数据表</param>
//        /// <param name="execlWorkBookStyle">execl单元格样式</param>
//        private static void InsertRow(DataTable dtSource, WorkBookStyle execlWorkBookStyle)
//        {
//            IWorkbook excelWorkbook = execlWorkBookStyle.BaseExcelWorkbook;
//            int rowCount = execlWorkBookStyle.TitleRowIndex.HasValue ? execlWorkBookStyle.TitleRowIndex.Value : 0;
//            //数据源导出数据集
//            ISheet newsheet = execlWorkBookStyle.CreateWorkSheet(execlWorkBookStyle.SheetName);
//            //警告区域
//            CreateAreaBlock(execlWorkBookStyle);
//            ICellStyle style = BorderCellStyle(excelWorkbook);
//            //设置表头
//            CreateHeader(newsheet, style, execlWorkBookStyle);

//            foreach (DataRow dr in dtSource.Rows)
//            {
//                rowCount++;
//                //创建行
//                IRow newRow = newsheet.CreateRow(rowCount);
//                //行数据添加
//                InsertCell(dtSource, dr, newRow, execlWorkBookStyle, rowCount);
//            }
//            //设置工作簿是否加锁
//            SetExcelSheetLock(execlWorkBookStyle.WorkBookIsLock, newsheet);
//            //当前工作薄添加数字限制区域
//            ExeclHelperExtensions.SheetAddValidationArea(newsheet, execlWorkBookStyle, ListColumnsName);
//            //设置选择下拉框
//            ExeclHelperExtensions.SetSelectValueArea(execlWorkBookStyle.ExeclColumnStyleList, ListColumnsName, newsheet);

//        }



//        /// <summary>
//        /// 导出数据行
//        /// </summary>
//        /// <param name="dtSource">数据表</param>
//        /// <param name="drSource">来源数据行</param>
//        /// <param name="currentExcelRow">当前数据行</param>
//        /// <param name="execlWorkBookStyle">表头样式</param>
//        /// <param name="rowCount">execl当前行</param>
//        private static void InsertCell(DataTable dtSource, DataRow drSource, IRow currentExcelRow, WorkBookStyle execlWorkBookStyle, int rowCount)
//        {
//            int cellIndex = 0;
//            foreach (var item in ListColumnsName)
//            {
//                //列名称
//                string columnsName = item.Key.ToString();
//                //根据列名称设置列样式
//                ExeclColumnStyle execlCellStyle = execlWorkBookStyle.ExeclColumnStyleList?.FirstOrDefault(it => it.ColumnsName == columnsName) ?? null;
//                ICellStyle columnStyle = execlWorkBookStyle.WorkSheet.GetColumnStyle(cellIndex);

//                //excelSheet.SetDefaultColumnStyle(cellIndex, SetExeclColumnStyle(columnStyle, excelWorkBook, execlCellStyle));
//                ICell newCell = null;
//                //获取字段类型
//                System.Type rowType = drSource[columnsName].GetType();
//                //获取字段值
//                string drValue = drSource[columnsName].ToString().Trim();
//                //错误消息列
//                string errColumn = columnsName + "ErrMsg";
//                string errMsg = string.Empty;
//                //是否有警告信息，有就设置错误批注
//                if (dtSource.Columns.Contains(errColumn))
//                {
//                    errMsg = drSource[errColumn].ToString().Trim();
//                }
//                switch (rowType.ToString())
//                {
//                    case "System.String"://字符串类型
//                        drValue = drValue.Replace("&", "&");
//                        drValue = drValue.Replace(">", ">");
//                        drValue = drValue.Replace("<", "<");
//                        newCell = currentExcelRow.CreateCell(cellIndex);
//                        newCell.SetCellValue(drValue);
//                        break;
//                    case "System.Guid"://GUID类型                        
//                        newCell = currentExcelRow.CreateCell(cellIndex);
//                        newCell.SetCellValue(drValue);
//                        break;
//                    case "System.DateTime"://日期类型
//                        DateTime dateV;
//                        DateTime.TryParse(drValue, out dateV);
//                        newCell = currentExcelRow.CreateCell(cellIndex);
//                        newCell.SetCellValue(dateV);
//                        //格式化显示
//                        ICellStyle cellStyle = execlWorkBookStyle.BaseExcelWorkbook.CreateCellStyle();
//                        IDataFormat format = execlWorkBookStyle.BaseExcelWorkbook.CreateDataFormat();
//                        cellStyle.DataFormat = format.GetFormat("yyyy-mm-dd hh:mm:ss");
//                        newCell.CellStyle = cellStyle;
//                        break;
//                    case "System.Boolean"://布尔型
//                        bool boolV = false;
//                        bool.TryParse(drValue, out boolV);
//                        newCell = currentExcelRow.CreateCell(cellIndex);
//                        newCell.SetCellValue(boolV);
//                        break;
//                    case "System.Int16"://整型
//                    case "System.Int32":
//                    case "System.Int64":
//                    case "System.Byte":
//                        int intV = 0;
//                        int.TryParse(drValue, out intV);
//                        newCell = currentExcelRow.CreateCell(cellIndex);
//                        //设置值
//                        newCell.SetCellValue(intV.ToString());
//                        break;
//                    case "System.Decimal"://浮点型
//                    case "System.Double":
//                        double doubV = 0;
//                        double.TryParse(drValue, out doubV);
//                        newCell = currentExcelRow.CreateCell(cellIndex);
//                        newCell.SetCellValue(doubV);
//                        break;
//                    case "System.DBNull"://空值处理
//                        newCell = currentExcelRow.CreateCell(cellIndex);
//                        newCell.SetCellValue("");
//                        break;
//                    default:
//                        throw (new Exception(rowType.ToString() + "：类型数据无法处理!"));
//                }
//                //设置列样式
//                if (execlCellStyle != null)
//                {
//                    SetColumnStyle(execlWorkBookStyle.WorkSheet, cellIndex, execlCellStyle);
//                    //设置单元格样式
//                    SetExeclCellStyle(newCell, execlCellStyle);
//                }
//                //设置批注
//                if (string.IsNullOrWhiteSpace(errMsg) == false)
//                {
//                    SetComment(errMsg, currentExcelRow, newCell, execlWorkBookStyle);
//                }
//                //设置计算表达式
//                SetCellFormula(execlCellStyle, execlWorkBookStyle.WorkSheet, newCell, rowCount);
//                cellIndex++;
//            }
//        }
//        #endregion

//        #region 单元格格式相关封装
//        /// <summary>
//        /// 设置单元格计算公式
//        /// </summary>
//        /// <param name="execlCellStyle">单元格样式</param>
//        /// <param name="excelSheet">execl工作薄</param>
//        /// <param name="newCell">单元格</param>
//        /// <param name="rowCount">第几行</param>
//        public static void SetCellFormula(ExeclColumnStyle execlCellStyle, ISheet excelSheet, ICell newCell, int rowCount)
//        {
//            if (execlCellStyle == null)
//            {
//                return;
//            }
//            //设置计算表达式
//            if (string.IsNullOrEmpty(execlCellStyle.SetCellFormula) == false)
//            {
//                int rowIndex = rowCount + 1;
//                //设置行的计算公式
//                newCell.SetCellFormula(string.Format(execlCellStyle.SetCellFormula, rowIndex));
//                if (excelSheet.ForceFormulaRecalculation == false)
//                {
//                    excelSheet.ForceFormulaRecalculation = true;//没有此句，则不会刷新出计算结果
//                }
//            }
//            //设置execl的单元格数据格式
//            if (string.IsNullOrEmpty(execlCellStyle.DataFormat) == false)
//            {
//                newCell.CellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat(execlCellStyle.DataFormat);
//            }
//        }
//        /// <summary>
//        /// 设置列的样式
//        /// </summary>
//        /// <param name="excelSheet">execl工作薄</param>
//        /// <param name="cellIndex">单元格位置</param>
//        /// <param name="execlCellStyle">列样式</param>
//        public static void SetColumnStyle(ISheet excelSheet, int cellIndex, ExeclColumnStyle execlCellStyle)
//        {
//            if (execlCellStyle == null)
//            {
//                return;
//            }
//            //设置列宽
//            if ((execlCellStyle?.Width ?? 0) > 0)
//            {
//                excelSheet.SetColumnWidth(cellIndex, execlCellStyle.Width);
//            }
//            //设置隐藏列
//            if (execlCellStyle?.IsHidden ?? false)
//            {
//                excelSheet.SetColumnHidden(cellIndex, true);
//            }
//        }
//        /// <summary>
//        /// 在execl工作表中创建指定的区块
//        /// </summary>
//        /// <param name="execlWorkBookStyle"></param>
//        public static void CreateAreaBlock(WorkBookStyle execlWorkBookStyle)
//        {
//            if (execlWorkBookStyle.AreaBlockList == null || execlWorkBookStyle.AreaBlockList.Count <= 0)
//            {
//                return;
//            }
//            foreach (AreaBlock item in execlWorkBookStyle.AreaBlockList)
//            {
//                //合并单元格
//                SetCellRangeAddress(execlWorkBookStyle.WorkSheet, item.Rowstart, item.Rowend, item.Colstart, item.Colend);
//                //创建第一行的说明列
//                IRow woringRow = execlWorkBookStyle.WorkSheet.CreateRow(item.Rowstart);
//                if (item.Height.HasValue)
//                {
//                    woringRow.Height = item.Height.Value;
//                }
//                ICell newCell = woringRow.CreateCell(item.Colstart);
//                newCell.CellStyle = item.CellStyle;
//                //警告文案
//                newCell.SetCellValue(item.Content);
//                newCell.CellStyle.WrapText = true;
//            }
//        }
//        /// <summary>
//        /// 设置下拉区域
//        /// </summary>
//        /// <param name="execlCellStyle"></param>
//        /// <param name="colIndex"></param>
//        /// <param name="excelSheet"></param>
//        public static void SetSelectValue(ExeclColumnStyle execlCellStyle, int colIndex, ISheet excelSheet)
//        {
//            //设置下拉区域
//            CellRangeAddressList regions = new CellRangeAddressList(execlCellStyle.StartBodyRow, execlCellStyle.EndBodyRow, colIndex, colIndex);
//            //设置下拉值
//            DVConstraint constraint = DVConstraint.CreateExplicitListConstraint(execlCellStyle.SelectValue);
//            HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
//            //不符合约束时的提示
//            dataValidate.CreateErrorBox("警告", "请从下拉框中选择");
//            //显示上面提示 = True    
//            dataValidate.ShowErrorBox = true;
//            excelSheet.AddValidationData(dataValidate);
//        }
//        /// <summary>
//        /// 工作簿加锁
//        /// </summary>
//        /// <param name="shouldLock">是否需要加锁</param>
//        /// <param name="excelSheet">工作薄</param>
//        public static void SetExcelSheetLock(bool shouldLock, ISheet excelSheet)
//        {

//            if (shouldLock == true)
//            {
//                excelSheet.ProtectSheet("mysoft");//设置密码保护
//            }
//        }

//        /// <summary>
//        /// 列锁定
//        /// </summary>
//        /// <param name="columnStyle"></param>
//        /// <param name="excelWorkBook"></param>
//        /// <param name="execlCellStyle"></param>
//        private static ICellStyle SetExeclColumnStyle(ICellStyle columnStyle, IWorkbook excelWorkBook, ExeclColumnStyle execlCellStyle)
//        {
//            if (columnStyle == null)
//            {
//                columnStyle = excelWorkBook.CreateCellStyle();
//            }
//            if (execlCellStyle != null)
//            {
//                columnStyle.IsLocked = execlCellStyle.IsLock;
//                columnStyle.IsHidden = execlCellStyle.IsHidden;
//            }
//            return columnStyle;
//        }
//        /// <summary>
//        /// 单元格格式
//        /// </summary>
//        /// <param name="execlCell"></param>
//        /// <param name="execlCellStyle"></param>
//        /// <returns></returns>
//        internal static void SetExeclCellStyle(ICell execlCell, ExeclColumnStyle execlCellStyle)
//        {
//            if (execlCell != null && execlCell.CellStyle != null && execlCellStyle != null)
//            {
//                execlCell.CellStyle = execlCellStyle.CellStyle;
//            }
//        }
//        /// <summary>
//        /// 设置批注
//        /// </summary>
//        /// <param name="errMsg"></param>
//        /// <param name="row"></param>
//        /// <param name="cell"></param>
//        /// <param name="execlWorkBookStyle"></param>
//        public static void SetComment(string errMsg, IRow row, ICell cell, WorkBookStyle execlWorkBookStyle)
//        {
//            //无错误信息不添加批注
//            if (string.IsNullOrEmpty(errMsg))
//            {
//                return;
//            }
//            ICreationHelper facktory = execlWorkBookStyle.BaseExcelWorkbook.GetCreationHelper();
//            HSSFPatriarch patr = (HSSFPatriarch)execlWorkBookStyle.WorkSheet.CreateDrawingPatriarch();
//            var anchor = facktory.CreateClientAnchor();
//            //设置批注区间大小
//            anchor.Col1 = cell.ColumnIndex;
//            anchor.Col2 = cell.ColumnIndex + 3;
//            //设置列
//            anchor.Row1 = row.RowNum;
//            anchor.Row2 = row.RowNum + 5;
//            //创建批注区域
//            var comment = patr.CreateCellComment(anchor);
//            comment.String = new HSSFRichTextString(errMsg);
//            //设置批注作者
//            comment.Author = ("mysoft");
//            cell.CellComment = (comment);
//        }
//        /// <summary>
//        /// 合并单元格
//        /// </summary>
//        /// <param name="sheet">要合并单元格所在的sheet</param>
//        /// <param name="rowstart">开始行的索引</param>
//        /// <param name="rowend">结束行的索引</param>
//        /// <param name="colstart">开始列的索引</param>
//        /// <param name="colend">结束列的索引</param>
//        public static void SetCellRangeAddress(ISheet sheet, int rowstart, int rowend, int colstart, int colend)
//        {
//            CellRangeAddress cellRangeAddress = new CellRangeAddress(rowstart, rowend, colstart, colend);
//            sheet.AddMergedRegion(cellRangeAddress);
//        }

//        /// <summary>
//        /// 添加导入时候的其他列
//        /// </summary>
//        public static void AddOtherColumn()
//        {

//        }




//        #endregion
//    }
//    /// <summary>
//    /// 实现排序接口，根据添加顺序导出
//    /// </summary>
//    public class NoSort : System.Collections.IComparer
//    {
//        /// <summary>
//        /// 自定义导出
//        /// </summary>
//        /// <param name="x"></param>
//        /// <param name="y"></param>
//        /// <returns></returns>
//        public int Compare(object x, object y)
//        {
//            return -1;
//        }
//    }
//}
