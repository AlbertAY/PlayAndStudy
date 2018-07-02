using ExeclTool.Model;
using NPOI.SS.UserModel;
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
            string temptPath = @"E:\github\PlayAndStudy\src\ExeclTool\TestFile\测试工作簿.xls";

            List<ISheet> sheetList;

            IWorkbook iWorkbook;

            using (FileStream fs = File.OpenRead(temptPath))
            {
                Tuple<IList<ExportModel>, IWorkbook, List<ISheet>> result = ExcelHelper.GetExeclDTO<ExportModel>(fs,0,0);

                iWorkbook = result.Item2;

                sheetList = result.Item3;
            }

            List<ExportModel> list = new List<ExportModel>();
            int max = 40;
            for (int i = 0; i < max; i++)
            {
                list.Add(new ExportModel()
                {
                    Name = "Albert" + 1,
                    Age = i
                });
            }
            //根据特性获取表头
            Dictionary<string, string> dicTitle = DataHelper.GetExeclTitleByTitleAttribute<ExportModel>();

            WorkBookStyle workBookStyle = new WorkBookStyle(dicTitle);
            workBookStyle.BaseExcelWorkbook = iWorkbook;

            workBookStyle.WorkSheet = sheetList.FirstOrDefault();

            MemoryStream memoryStream = StyleFromTemplateExeclHelper.ExportExcel(list, workBookStyle);

            //MemoryStream memoryStream = ProductImportExeclHelper.ExportExcel(list, workBookStyle);

            SaveToFile(memoryStream,DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss-ffff") +".xls");


        }
        /// <summary>
        /// 保存文件到硬盘上面
        /// </summary>
        /// <param name="ms"></param>
        /// <param name="fullPath"></param>
        public static void SaveToFile(MemoryStream ms, string fileName)
        {
            string savePath = @"E:\github\PlayAndStudy\src\ExeclTool\TestFile\Result";
            string fullPath = Path.Combine(savePath,fileName);
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
        [ExeclTitle("姓名")]
        public string Name { set; get; }

        [ExeclTitle("年龄")]
        public int Age { get; set; }

    }



}
