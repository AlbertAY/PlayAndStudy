﻿using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExeclTool.Model;

namespace ExeclTool
{
    /// <summary>
    /// execl帮助扩展
    /// </summary>
    public static class ExeclHelperExtend
    {
        /// <summary>
        /// 金额可编辑区域
        /// </summary>
        /// <param name="execlCellStyleEntity"></param>
        /// <param name="excelWorkbook"></param>
        public static void SetDecimalEditStyle(ExeclColumnStyle execlCellStyleEntity, IWorkbook excelWorkbook)
        {
            execlCellStyleEntity.Width = 15;
            //编辑背景为白色
            execlCellStyleEntity.BackGrandIndexed = NPOI.HSSF.Util.HSSFColor.White.Index;
            //表头样式
            execlCellStyleEntity.TitleStyle = ExeclHelperExtend.EditTitleStyle(excelWorkbook);
            //该字段为编辑字段，设置为编辑样式
            execlCellStyleEntity.CellStyle = ExeclHelperExtend.EditStyle(excelWorkbook);
            //金额右对齐
            execlCellStyleEntity.CellStyle.Alignment = HorizontalAlignment.Right;
            execlCellStyleEntity.ValidationType = ValidationType.DECIMAL.ToString();
        }
        /// <summary>
        /// 非编辑列头部样式
        /// </summary>
        /// <param name="excelWorkbook"></param>
        /// <returns></returns>
        public static ICellStyle NoEditTitleStyle(IWorkbook excelWorkbook)
        {
            ICellStyle style = excelWorkbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(GetTitleFont(excelWorkbook));
            //设置背景色
            style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
            style.FillPattern = FillPattern.SolidForeground;
            return style;
        }
        /// <summary>
        /// 编辑列头部样式
        /// </summary>
        /// <param name="excelWorkbook"></param>
        /// <returns></returns>
        public static ICellStyle EditTitleStyle(IWorkbook excelWorkbook)
        {
            ICellStyle style = excelWorkbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.Alignment = HorizontalAlignment.Center;
            style.SetFont(GetTitleFont(excelWorkbook));
            //设置背景色
            style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightOrange.Index;
            style.FillPattern = FillPattern.SolidForeground;
            return style;
        }
        /// <summary>
        /// 编辑单元格样式
        /// </summary>
        /// <param name="excelWorkbook"></param>
        /// <returns></returns>
        public static ICellStyle NoEditStyle(IWorkbook excelWorkbook)
        {
            ICellStyle style = excelWorkbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.Alignment = HorizontalAlignment.Left;
            style.SetFont(GetEditFont(excelWorkbook));
            //设置背景色
            style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
            style.FillPattern = FillPattern.SolidForeground;
            return style;
        }
        /// <summary>
        /// 非编辑单元格样式
        /// </summary>
        /// <param name="excelWorkbook"></param>
        /// <returns></returns>
        public static ICellStyle EditStyle(IWorkbook excelWorkbook)
        {
            ICellStyle style = excelWorkbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.Alignment = HorizontalAlignment.Left;
            style.SetFont(GetEditFont(excelWorkbook));
            return style;
        }
        /// <summary>
        /// 头部字体
        /// </summary>
        /// <param name="excelWorkbook"></param>
        /// <returns></returns>
        public static IFont GetTitleFont(IWorkbook excelWorkbook)
        {
            HSSFFont font = (HSSFFont)excelWorkbook.CreateFont();
            font.FontHeightInPoints = 10;//字号 
            font.FontName = "微软雅黑";
            font.Color = NPOI.HSSF.Util.HSSFColor.Black.Index;//颜色 
            font.Boldweight = 600;//HSSFFont.BOLDWEIGHT_BOLD;//加粗  

            return font;
        }

        /// <summary>
        /// 编辑单元格字体
        /// </summary>
        /// <param name="excelWorkbook"></param>
        /// <returns></returns>
        public static IFont GetEditFont(IWorkbook excelWorkbook)
        {
            HSSFFont font = (HSSFFont)excelWorkbook.CreateFont();
            font.FontHeightInPoints = 10;//字号 
            font.FontName = "微软雅黑";
            font.Color = NPOI.HSSF.Util.HSSFColor.Black.Index;//颜色 
            return font;
        }

        /// <summary>
        /// 添加区块
        /// </summary>
        /// <param name="workBookStyle"></param>
        public static void AddAreaBlock(WorkBookStyle workBookStyle)
        {
            //创建区块区域
            workBookStyle.AreaBlockList = new List<AreaBlock>();
            //获取单元格样式
            ICellStyle cellStyle = ExeclHelperExtend.EditStyle(workBookStyle.BaseExcelWorkbook);
            cellStyle.Alignment = HorizontalAlignment.Left;
            cellStyle.WrapText = true;
            //区块的区域
            AreaBlock areaBlock = new AreaBlock()
            {
                Rowstart = 0,
                Rowend = 0,
                Colstart = 1,
                Colend = 6,
                Height = 30 * 30,
                Content = "警告   \r\n 1：测试警告",
                CellStyle = cellStyle
            };
            workBookStyle.AreaBlockList.Add(areaBlock);
        }
        /// <summary>
        /// 获取警告语
        /// </summary>
        /// <returns></returns>
        public static string GetWaring()
        {
            string msg = "";
            string[] str = msg.Split(new char[] { '.' });
            return string.Join("\n", str);//数组转成字符串 
        }
        /// <summary>
        /// 添加验证区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="execlWorkBookStyle"></param>
        /// <param name="listColumnsName"></param>
        public static void SheetAddValidationArea(ISheet sheet, WorkBookStyle execlWorkBookStyle, Dictionary<string, string> listColumnsName)
        {
            //列限制区域
            IEnumerable <ExeclColumnStyle> setValueCellList = execlWorkBookStyle.ExeclColumnStyleList.Where(it => string.IsNullOrEmpty(it.ValidationType) == false);


            string[] columnNames = listColumnsName.Select(it => it.Key).ToArray();
            foreach (ExeclColumnStyle item in setValueCellList)
            {
                for (int i = 0; i < columnNames.Length; i++)
                {
                    if (item.ColumnsName.Equals(columnNames[i]))
                    {
                        SetCellInputNumber(sheet, item.StartBodyRow, item.EndBodyRow, i, i);
                        continue;
                    }
                }
            }
        }
        /// <summary>
        /// 设置下拉选择范围
        /// </summary>
        /// <param name="execlCellStyleList"></param>
        /// <param name="listColumnsName"></param>
        /// <param name="excelSheet"></param>
        public static void SetSelectValueArea(List<ExeclColumnStyle> execlCellStyleList, Dictionary<string, string> listColumnsName, ISheet excelSheet)
        {
            IEnumerable<ExeclColumnStyle> setValueCellList = execlCellStyleList.Where(it => it.SelectValue != null && it.SelectValue.Length > 0);
            string[] columnNames = listColumnsName.Select(it => it.Key).ToArray();
            foreach (ExeclColumnStyle item in setValueCellList)
            {
                for (int i = 0; i < columnNames.Length; i++)
                {
                    if (item.ColumnsName.Equals(columnNames[i]))
                    {
                        ExcelHelper.SetSelectValue(item, i, excelSheet);
                        continue;
                    }
                }
            }
        }
        /// <summary>
        /// 设置只是输入数字
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="firstRow"></param>
        /// <param name="lastRow"></param>
        /// <param name="firstCol"></param>
        /// <param name="lastCol"></param>
        public static void SetCellInputNumber(ISheet sheet, int firstRow, int lastRow, int firstCol, int lastCol)
        {
            var cellRegions = new CellRangeAddressList(firstRow, lastRow, firstCol, firstCol);
            DVConstraint constraint = DVConstraint.CreateNumericConstraint(ValidationType.DECIMAL, OperatorType.BETWEEN, "0", "999999999");
            HSSFDataValidation dataValidate = new HSSFDataValidation(cellRegions, constraint);
            dataValidate.CreateErrorBox("警告", "请输入数字");
            sheet.AddValidationData(dataValidate);
        }
    }
}