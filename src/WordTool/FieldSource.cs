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
        }

        public Guid Oid { set; get; }


        public List<FieldSource> FieldSourceList { set; get; }


        public int 



    }
    public class FieldSource
    {
        public FieldSource(string sql)
        {
            SourceSql = sql;
        }
        /// <summary>
        /// 来源sql
        /// </summary>
        public string SourceSql { set; get; }

        public string SoucreType { set; get; }

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
