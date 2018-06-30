using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExeclTool
{
    /// <summary>
    /// 集合转化帮助
    /// </summary>
    public static class DataHelper
    {

        #region 将Datatable转换成List泛型

        /// <summary>
        /// table转业务对象
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="table"></param>
        /// <returns></returns>
        public static IList<T> ModelConver<T>(this DataTable table)
        {
            // 定义集合    
            IList<T> entityList = new List<T>();

            // 获得此模型的类型   
            Type type = typeof(T);
            foreach (DataRow dr in table.Rows)
            {
                T entity = System.Activator.CreateInstance<T>();
                // 获得此模型的公共属性      
                PropertyInfo[] propertys = entity.GetType().GetProperties();

                //标准表头：实体定义的
                List<string> standardHead = new List<string>();

                //错误信息
                List<ErrorMessage> errorMessages = new List<ErrorMessage>();
                foreach (PropertyInfo propertyInfo in propertys)
                {
                    //判断是否有这个特性
                    ExeclTitleAttribute execlTitleAttribute = propertyInfo.GetCustomAttribute<ExeclTitleAttribute>();
                    if (execlTitleAttribute == null)
                    {
                        continue;
                    }
                    if (table.Columns.Contains(execlTitleAttribute.TitleName))
                    {
                        if (execlTitleAttribute.ColumnIndex == null)
                        {
                            //得到该列的位置索引
                            int index = table.Columns.IndexOf(execlTitleAttribute.TitleName);
                            execlTitleAttribute.ColumnIndex = index;
                        }
                        standardHead.Add(execlTitleAttribute.TitleName);
                        // 判断此属性是否有Setter      
                        if (propertyInfo.CanWrite == false)
                        {
                            continue;
                        }
                        object value = dr[execlTitleAttribute.TitleName];
                        if (value != null)
                        {
                            execlTitleAttribute.ErrorMessage = propertyInfo.SetPropertyInfoValue(entity, value);
                            //将错误信息添加到错误集合
                            if (string.IsNullOrEmpty(execlTitleAttribute.ErrorMessage) == false)
                            {
                                ErrorMessage err = new ErrorMessage
                                {
                                    Name = propertyInfo.Name,
                                    ErrMsg = execlTitleAttribute.ErrorMessage
                                };
                                errorMessages.Add(err);
                            }
                        }
                    }
                }

                #region 设置异常列信息
                PropertyInfo errorMessagePropertyInfo = entity.GetType().GetProperty("ErrorMessages");
                if (errorMessagePropertyInfo != null && errorMessages.Count > 0)
                {
                    errorMessagePropertyInfo.SetValue(entity, errorMessages, null);
                }
                #endregion

                AddOtherColumns(table.Columns, standardHead, dr, entity);
                entityList.Add(entity);
            }
            return entityList;
        }
        /// <summary>
        /// 获取对象的列
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static Dictionary<string, string> GetExeclTitleByTitleAttribute<T>()
        {
            List<ExeclTitleAttribute> list = new List<ExeclTitleAttribute>();
            //取属性上的自定义特性
            Type objType = typeof(T);
            //遍历对象的所有属性
            foreach (PropertyInfo propInfo in objType.GetProperties())
            {
                //得到指定的特性
                ExeclTitleAttribute attr = propInfo.GetCustomAttribute<ExeclTitleAttribute>();
                if (attr != null)
                {
                    if (attr.ColumnIndex.HasValue == false)
                    {
                        attr.ColumnIndex = 0;
                    }
                    attr.PropName = propInfo.Name;
                    list.Add(attr);
                }
            }
            //按照排序字段进行排序
            list = list.OrderBy(o => o.ColumnIndex).ToList();
            //转成表头字典并且返回
            return list.ToDictionary(key => key.PropName, value => value.TitleName);
        }

        /// <summary>
        /// 添加非标准列的值到行对象其他列（OtherColumns）中
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataColumnCollection"></param>
        /// <param name="standardHead"></param>
        /// <param name="dataRow"></param>
        /// <param name="model"></param>
        public static void AddOtherColumns<T>(DataColumnCollection dataColumnCollection, List<string> standardHead, DataRow dataRow, T model)
        {
            //定义其他列，即Excel列不在标准的实体里面定义的
            PropertyInfo propertyInfo = model.GetType().GetProperty("OtherColumns");
            if (propertyInfo == null)
            {
                return;
            }

            //其他列字典，如：key:材料名称，value:格力空调
            List<ColumnInfo> otherColumns = new List<ColumnInfo>();
            foreach (DataColumn item in dataColumnCollection)
            {
                //判读如果列不在标准的头部里面则加入记录
                string columnName = item.ColumnName;
                if (standardHead.Contains(columnName) == false)
                {
                    ColumnInfo columnInfo = new ColumnInfo()
                    {
                        Title = columnName,
                        Value = dataRow[columnName].ToString(),
                    };
                    otherColumns.Add(columnInfo);
                }
            }
            if (otherColumns.Count <= 0)
            {
                return;
            }
            //设置属性值
            propertyInfo.SetValue(model, otherColumns, null);
        }

        /// <summary>
        /// 设置属性值，如果失败则范围异常信息
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="propertyInfo"></param>
        /// <param name="model"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string SetPropertyInfoValue<T>(this PropertyInfo propertyInfo, T model, object value)
        {
            //对象mapping的转string
            if (propertyInfo.PropertyType == typeof(string))
            {
                propertyInfo.SetValue(model, value, null);
            }
            //对象mapping的转Int
            else if (propertyInfo.PropertyType == typeof(int))
            {
                int valueResult;
                if (int.TryParse(value.ToString(), out valueResult))
                {
                    propertyInfo.SetValue(model, valueResult, null);
                }
                else
                {
                    return "请输入有效的数字";
                }
            }
            //对象mapping的转Int?
            else if (propertyInfo.PropertyType == typeof(int?))
            {
                int resultValue;
                if (int.TryParse(value.ToString(), out resultValue))
                {
                    propertyInfo.SetValue(model, resultValue, null);
                }
                else
                {
                    propertyInfo.SetValue(model, null, null);
                }
            }
            //对象mapping的转时间
            else if (propertyInfo.PropertyType == typeof(DateTime))
            {
                DateTime valueResult;
                if (DateTime.TryParse(value.ToString(), out valueResult))
                {
                    propertyInfo.SetValue(model, valueResult, null);
                }
                else
                {
                    return "请输入有效的时间类型";
                }
            }
            //对象mapping的转时间，可为空类型
            else if (propertyInfo.PropertyType == typeof(DateTime?))
            {
                DateTime resultValue;
                if (DateTime.TryParse(value.ToString(), out resultValue))
                {
                    propertyInfo.SetValue(model, resultValue, null);
                }
                else
                {
                    propertyInfo.SetValue(model, null, null);
                }
            }
            //对象mapping的转decimal
            else if (propertyInfo.PropertyType == typeof(decimal))
            {
                decimal resultValue;
                if (decimal.TryParse(value.ToString(), out resultValue))
                {
                    propertyInfo.SetValue(model, resultValue, null);
                }
                else
                {
                    return "请输入有效的数字";
                }
            }
            //对象mapping的转decimal，可为空类型
            else if (propertyInfo.PropertyType == typeof(decimal?))
            {
                decimal resultValue;
                if (decimal.TryParse(value.ToString(), out resultValue))
                {
                    propertyInfo.SetValue(model, resultValue, null);
                }
                else
                {
                    propertyInfo.SetValue(model, null, null);
                }
            }
            //对象mapping的转Guid
            else if (propertyInfo.PropertyType == typeof(Guid))
            {
                Guid valueResult;
                if (Guid.TryParse(value.ToString(), out valueResult))
                {
                    propertyInfo.SetValue(model, valueResult, null);
                }
                else
                {
                    return "请输入有效的GUID";
                }
            }
            //对象mapping的转Guid类型
            else if (propertyInfo.PropertyType == typeof(Guid?))
            {
                Guid resultValue;
                if (Guid.TryParse(value.ToString(), out resultValue))
                {
                    propertyInfo.SetValue(model, resultValue, null);
                }
                else
                {
                    propertyInfo.SetValue(model, null, null);
                }
            }
            return string.Empty;
        }

        #endregion

        #region 将Datatable转换成List泛型
        /// <summary>
        /// 将数据表转成泛型集合
        /// </summary>
        /// <typeparam name="T">类型</typeparam>
        /// <param name="table">数据表</param>
        /// <returns></returns>
        public static IList<T> ConvertTo<T>(this DataTable table)
        {
            if (table == null)
            {
                return null;
            }
            //创建对象集合
            List<DataRow> rows = new List<DataRow>();
            foreach (DataRow row in table.Rows)
            {
                //添加行
                rows.Add(row);
            }
            //转成对象集合
            return ConvertTo<T>(rows);
        }
        /// <summary>
        /// 数据行集合转成泛型集合
        /// </summary>
        /// <typeparam name="T">泛型类型</typeparam>
        /// <param name="rows">行集合</param>
        /// <returns></returns>
        private static IList<T> ConvertTo<T>(IList<DataRow> rows)
        {
            IList<T> list = null;
            //行不为空
            if (rows != null)
            {
                //创建泛型集合
                list = new List<T>();
                foreach (DataRow row in rows)
                {
                    //创建对象
                    T item = CreateItem<T>(row);
                    list.Add(item);
                }
            }
            return list;
        }
        /// <summary>
        /// 数据行转成泛型对象
        /// </summary>
        /// <typeparam name="T">需要转成的类型</typeparam>
        /// <param name="row">数据行</param>
        /// <returns></returns>
        private static T CreateItem<T>(DataRow row)
        {
            T obj = default(T);
            if (row != null)
            {
                //创建实例
                obj = Activator.CreateInstance<T>();
                foreach (DataColumn column in row.Table.Columns)
                {
                    //获取字段属性
                    PropertyInfo prop = obj.GetType().GetProperty(column.ColumnName);
                    try
                    {
                        object value = row[column.ColumnName];
                        //赋值
                        if (value != System.DBNull.Value)
                        {
                            //添加值
                            prop.SetValue(obj, value, null);
                        }                           
                    }
                    catch
                    {
                        //异常直接抛出
                        throw;
                    }
                }
            }
            return obj;
        }
        #endregion

        #region 将List转换成Datatable
        /// <summary>
        /// 将泛型集合转成数据表
        /// </summary>
        /// <param name="list">数据集合</param>
        /// <returns></returns>
        public static DataTable ToDataTable(IList list)
        {
            //创建数据表对象
            DataTable result = new DataTable();
            if (list.Count > 0)
            {
                //获取属性集合
                PropertyInfo[] propertys = list[0].GetType().GetProperties();
                foreach (PropertyInfo pi in propertys)
                {
                    try
                    {
                        Type colType = pi.PropertyType;
                        //为空的值类型的特殊处理
                        if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition() == typeof(Nullable<>)))
                        {
                            colType = colType.GetGenericArguments()[0];
                        }
                        //添加到列
                        result.Columns.Add(pi.Name, colType);
                    }
                    catch (Exception ex)
                    {
                        //抛出异常
                        throw ex;
                    }
                    
                }
                for (int i = 0; i < list.Count; i++)
                {
                    //创建数组
                    ArrayList tempList = new ArrayList();
                    //根据属性分别往数组里面添加值
                    foreach (PropertyInfo pi in propertys)
                    {
                        object obj = pi.GetValue(list[i], null);                    
                        tempList.Add(obj);
                    }
                    object[] array = tempList.ToArray();
                    //添加值
                    result.LoadDataRow(array, true);
                }
            }
            return result;
        }

        #endregion
        /// <summary>
        /// 生成ID
        /// </summary>
        /// <returns></returns>
        public static string GenerateStringID()
        {
            long i = 1;
            //生成新的id
            foreach (byte b in Guid.NewGuid().ToByteArray())
            {
                i *= ((int)b + 1);
            }
            //返回ID
            return string.Format("{0:x16}", i - DateTime.Now.Ticks).Substring(0, 11);
        }
    }
}
