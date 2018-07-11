using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XWPF.UserModel;

namespace WordTool
{
    public class WordExport
    {


        public void WordStart(string path)
        {

            using (FileStream fs = File.OpenRead(path))
            {
               

                XWPFDocument doc = new XWPFDocument(fs);     
                FileChange(doc);
                MemoryStream memoryStream = new MemoryStream();
                doc.Write(memoryStream);                      
                memoryStream.Close();
                SaveToFile(memoryStream,$"11111-{DateTime.Now.ToString("yyyyMMdd-HHmmssffff")}.doc");

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
                        XWPFTableCell rowCell= xWPFTableRow.GetCell(j);
                        if (j==0)
                        {
                            rowCell.SetText("张三" + i);
                        }
                        else
                        {
                            rowCell.SetText(i.ToString());
                        }
                    }
                }


                //foreach (XWPFTableCell cell in row.GetTableCells())
                //{

                //}




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

    }
}
