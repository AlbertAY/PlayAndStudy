using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XWPF.UserModel;
using Dapper;
using System.Data;
using NPOI.SS.Formula.Functions;
using NPOI.HWPF;
using NPOI;

namespace WordTool
{
    public class WordExport
    {

        public void TestExport(string path)
        {
            WorldField worldField = new WorldField(new Guid("14ca7877-d988-e811-95e9-94c691050e22"));


            string sql = @"SELECT   ContractCode [合同编码] ,
        ContractName [合同名称] ,
        JbrName [合同经办人] ,
        JfProviderName 甲方单位,
        YfProviderName 乙方单位,
        SignDate 签约日期,
        FormationMode 采购方式,
        StrategicAgreement 战略协议,
        DeptName 经办部门,
        JbrName 经办人,
        TotalJsAmount 合同结算金额,
        HtAmount [合同金额（含税）],
        HtNonTaxAmount  [合同金额（不含税）],
        HtInputTaxAmount [合同进项税额],
        ISNULL(AdjustAmount_Bz, 0.00) [调整后合同金额],
        ISNULL(TotalApplyAmount_Bz, 0.00) 累计申请金额 ,
        ISNULL(TotalPaidAmount, 0.00)  累计实付金额,
        ISNULL(TotalDeductAmount_Bz, 0.00) 累计应扣金额,
        ProjectName 项目名称,
        HtTypeName 合同类型,
        PriceRatioAvg [加成率（综合）],
        TotalInvoiceAmount 累计发票金额,
        ISNULL(AdjustTotalAmount, 0.00) AdjustTotalAmount
FROM   dbo.vcl_Contract
                            WHERE   ContractGUID={oid}";

            FieldSource contract = new FieldSource(sql, "合同信息", SoucreType.Entity);


            string productList = @"SELECT  p.Code [材料编码],
        p.Name [材料名称],
        p.BrandName [品牌名称],
        p.Model [材料型号],
        p.Attribute [指标属性],
        p.Unit 单位,
        cd.Price [单价（含税）],
        cd.Count [数量],
        cd.StrategicAgreementPrice 协议单价,
        cd.NotTaxPrice [单价（不含税）],
        cd.NoTaxAmount [金额（不含税）],
        cd.Amount [金额（含税）],
        cd.TaxAmount 税额 ,
        cd.TaxRate 税率,
        cd.Remark 备注
FROM dbo.cl_ContractProductDetails cd
        INNER JOIN dbo.cl_Product p ON cd.ProductGUID = p.ProductGUID
                            WHERE ContractGUID = {oid}";

            FieldSource contractProduct = new FieldSource(productList, "合同材料", SoucreType.List);

            worldField.FieldSourceList.Add(contract);

            worldField.FieldSourceList.Add(contractProduct);

            using (FileStream fs = File.OpenRead(path))
            {

                XWPFDocument doc;
                try
                {
                     doc = new XWPFDocument(fs);
                }
                catch (Exception)
                {
                    //doc = new HWPFDocument(fs);
                    throw;
                }                
                Begin(worldField, doc);
                MemoryStream memoryStream = new MemoryStream();
                doc.Write(memoryStream);
                memoryStream.Close();
                SaveToFile(memoryStream, $"11111-{DateTime.Now.ToString("yyyyMMdd-HHmmssffff")}.doc");

            }
        }

        public void Begin(WorldField worldField, XWPFDocument doc)
        {

            if (worldField.FieldSourceList == null || worldField.FieldSourceList.Count <= 0)
            {
                throw new Exception("请先配置导出数据源");
            }
            var modelField = worldField.FieldSourceList.Where(it => it.SoucreType == SoucreType.Entity);

            if ((modelField?.Count() ?? 0) > 1)
            {
                throw new Exception("只能配置一个主体表信息");
            }
            //主体的字段信息替换
            if (modelField != null)
            {
                var baseModel = modelField.FirstOrDefault();
                string executSql = baseModel.SourceSql.Replace("{oid}", $"'{worldField.Oid}'");
                DataTable dataTable = new DataTable();
                using (SqlConnection conn = GetSqlConnection())
                {
                    dataTable.Load(conn.ExecuteReader(executSql));
                }
                if (dataTable.Rows.Count <= 0)
                {
                    throw new Exception("主题信息查询失败");
                }

                DataRow row = dataTable.Rows[0];

                List<Field> fieldList = GetFieldList(dataTable.Columns, "");

                IList<XWPFParagraph> paragraphs = doc.Paragraphs;
                foreach (XWPFParagraph xwpParagraph in paragraphs)
                {


                    foreach (XWPFRun xwprun in xwpParagraph.Runs)
                    {
                        string text = xwprun.Text.Clone().ToString() ;
                        if (text.Contains("加成率"))
                        {
                            
                        }
                        foreach (Field field in fieldList)
                        {
                            if (xwprun.Text.Contains(field.SignName))
                            {
                                text = text.Replace(field.SignName, row[field.Name].ToString());                               
                            }
                        }
                        xwprun.ReplaceText(xwprun.Text, text);

                    }
                }
            }
            //可重复表数量
            List<FieldSource> tableSource = worldField.FieldSourceList.Where(it => it.SoucreType == SoucreType.List)?.ToList() ?? new List<FieldSource>();


            IList<XWPFTable> tables = doc.Tables;


            foreach (FieldSource item in tableSource)
            {
                string executSql = item.SourceSql.Replace("{oid}", $"'{worldField.Oid}'");

                DataTable dataTable = new DataTable();
                using (SqlConnection conn = GetSqlConnection())
                {
                    dataTable.Load(conn.ExecuteReader(executSql));
                }
                if (dataTable.Rows.Count <= 0)
                {
                    throw new Exception("数据查询失败");
                }
                List<Field> fieldList = GetFieldList(dataTable.Columns, item.TableName);

                //具体的表处理
                foreach (XWPFTable table in tables)
                {
                    XWPFTableRow row = table.Rows.FirstOrDefault();

                    Dictionary<int, string> dicTitle = GetTableField(row, item.TableName, fieldList);
                    if (dicTitle.Count <= 0)
                    {
                        continue;
                    }

                    foreach (DataRow dataRow in dataTable.Rows)
                    {
                        XWPFTableRow xWPFTableRow = table.CreateRow();
                        foreach (var title in dicTitle)
                        {
                            string columnName = fieldList.FirstOrDefault(it => it.SignName.Equals(title.Value))?.Name;
                            if (string.IsNullOrEmpty(columnName))
                            {
                                continue;
                            }
                            XWPFTableCell rowCell = xWPFTableRow.GetCell(title.Key);
                            rowCell.SetText(dataRow[columnName].ToString());
                        }
                    }
                }
            }

        }
        /// <summary>
        /// 获取表头的名称
        /// </summary>
        /// <param name="table"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public Dictionary<int, string> GetTableField(XWPFTableRow row, string tableName, List<Field> fieldList)
        {
            Dictionary<int, string> tableField = new Dictionary<int, string>();

            for (int j = 0; j < row.GetTableCells().Count; j++)
            {
                string text = row.GetCell(j).GetText();
                if (text.Contains(tableName))
                {
                    tableField.Add(j, text);
                    Field field = fieldList.FirstOrDefault(it => it.SignName == text);
                    if (field != null)
                    {
                        //移除第一行数据
                        row.GetCell(j).RemoveParagraph(0);
                        row.GetCell(j).SetText(field.Name);
                    }
                }
            }
            return tableField;
        }




        public List<Field> GetFieldList(DataColumnCollection columnCollections, string tableName)
        {
            string sign = "｛{0}｝";
            if (string.IsNullOrEmpty(tableName) == false)
            {
                sign = "｛" + tableName + "-{0}｝";
            }

            List<Field> fieldList = new List<Field>();
            foreach (DataColumn item in columnCollections)
            {
                Field entity = new Field();

                string signName = string.Format(sign, item.ColumnName);
                entity.Name = item.ColumnName;
                entity.SignName = signName;
                fieldList.Add(entity);
            }
            return fieldList;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public SqlConnection GetSqlConnection()
        {
            SqlConnection conn = new SqlConnection(@"Data Source=10.5.6.9\sql2014;Initial Catalog=dotnet_erp60_clgyl_bigdata;User ID=sa;Password=95938");
            conn.Open();
            return conn;
        }




        #region old test
        public void WordStart(string path)
        {

            using (FileStream fs = File.OpenRead(path))
            {


                XWPFDocument doc = new XWPFDocument(fs);
                FileChange(doc);
                MemoryStream memoryStream = new MemoryStream();
                doc.Write(memoryStream);
                memoryStream.Close();
                SaveToFile(memoryStream, $"11111-{DateTime.Now.ToString("yyyyMMdd-HHmmssffff")}.doc");

            }
        }

        private void FileChange(XWPFDocument doc)
        {
            string title = @"｛合同名称｝";
            IList<XWPFParagraph> paragraphs = doc.Paragraphs;
            foreach (XWPFParagraph item in paragraphs)
            {

                foreach (XWPFRun xwprun in item.Runs)
                {
                    if (xwprun.Text.Contains(title))
                    {
                        xwprun.ReplaceText(title, "艾勇的合同");
                    }
                }
            }
            IList<XWPFTable> tables = doc.Tables;
            foreach (XWPFTable table in tables)
            {

                XWPFTableRow row = table.Rows.FirstOrDefault();

                for (int i = 0; i < 10; i++)
                {
                    XWPFTableRow xWPFTableRow = table.CreateRow();
                    for (int j = 0; j < row.GetTableCells().Count; j++)
                    {
                        XWPFTableCell rowCell = xWPFTableRow.GetCell(j);
                        if (j == 0)
                        {
                            rowCell.SetText("张三" + i);
                        }
                        else
                        {
                            rowCell.SetText(i.ToString());
                        }
                    }
                }

            }




        }

        /// <summary>
        /// 保存文件到硬盘上面
        /// </summary>
        /// <param name="ms"></param>
        /// <param name="fullPath"></param>
        public static void SaveToFile(MemoryStream ms, string fileName)
        {
            string savePath = @"E:\github\PlayAndStudy\src\WordTool\Result";
            string fullPath = Path.Combine(savePath, fileName);
            string dirName = Path.GetDirectoryName(fullPath);
            if (!Directory.Exists(dirName))//判断是否存在
            {
                Directory.CreateDirectory(dirName);//创建新路径
            }
            using (FileStream fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write))
            {
                byte[] data = ms.ToArray();
                fs.Write(data, 0, data.Length);
                fs.Flush();
                data = null;
            }
        }
        #endregion

    }
}
