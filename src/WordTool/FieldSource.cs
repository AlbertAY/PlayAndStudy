using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordTool
{

    public class WorldField
    {
        /// <summary>
        /// 当前需要打印的业务的主键
        /// </summary>
        /// <param name="oid"></param>
        public WorldField(Guid oid)
        {
            Oid = oid;
            FieldSourceList = new List<FieldSource>();
        }

        public Guid Oid { set; get; }


        public List<FieldSource> FieldSourceList { set; get; }

    }
    public class FieldSource
    {
        public FieldSource(string sql, SoucreType soucreType)
        {
            SourceSql = sql;
            SoucreType = soucreType;
        }
        public FieldSource(string sql, string tableName, SoucreType soucreType)
        {
            SourceSql = sql;
            SoucreType = soucreType;
            TableName = tableName;
        }
        /// <summary>
        /// 来源sql
        /// </summary>
        public string SourceSql { set; get; }

        public SoucreType SoucreType { set; get; }

        public string TableName { set; get; }

    }


    public class Field
    {
        /// <summary>
        /// 名称
        /// </summary>
        public string Name { set; get; }
        /// <summary>
        /// 标识名称
        /// </summary>

        public string SignName { set; get; }
    }

    public enum SoucreType
    {
        /// <summary>
        /// 主体信息，比如合同信息
        /// </summary>
        Entity=1,
        /// <summary>
        /// 先关列表，比如合同下的材料明细
        /// </summary>
        List=2
    }
}
