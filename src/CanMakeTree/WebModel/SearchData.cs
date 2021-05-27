using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CanMakeTree.WebModel
{
    public class SearchData
    {
        //分子Id
        public int MoleculeId { set; get; }

        //分子的smiles
        public string Smiles { set; get; }

        //分子结果
        public string Result { set; get; }
    }
}
