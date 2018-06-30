using ExeclTool.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExeclTool
{
    class Program
    {
        static void Main(string[] args)
        {
            List<ExportModel> list = new List<ExportModel>();
            int max = 20;
            for (int i = 0; i < max; i++)
            {
                list.Add(new ExportModel()
                {
                    Name = "Albert" + 1,
                    Age = i
                });
            }
            Dictionary<string, string> dicTitle = new Dictionary<string, string>();
            dicTitle.Add("Name", "姓名");
            dicTitle.Add("Age", "年龄");
            WorkBookStyle workBookStyle = new WorkBookStyle(dicTitle);
            ExcelHelper.ExportExcel(list, workBookStyle,"测试导出");
        }
    }



    public class ExportModel
    {

        public string Name { set; get; }


        public int Age { get; set; }

    }



}
