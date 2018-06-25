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
