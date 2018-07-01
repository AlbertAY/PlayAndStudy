using ExeclTool.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExeclTool
{
    class Program
    {
        static void Main(string[] args)
        {
            string temptPath = @"C:\Users\Albert\Desktop\测试导入.xlsx";
            using (FileStream fs = File.OpenRead(temptPath))
            {

            }

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
            MemoryStream memoryStream = ProductImportExeclHelper.ExportExcel(list, workBookStyle);

            string savePath = @"E:\测试文件";

            SaveToFile(memoryStream,savePath);


        }
        /// <summary>
        /// 保存文件到硬盘上面
        /// </summary>
        /// <param name="ms"></param>
        /// <param name="fullPath"></param>
        public static void SaveToFile(MemoryStream ms, string fullPath)
        {
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
    }



    public class ExportModel
    {

        public string Name { set; get; }


        public int Age { get; set; }

    }



}
