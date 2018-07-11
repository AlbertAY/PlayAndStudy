using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordTool
{
    class Program
    {
        static void Main(string[] args)
        {
            string temptPath = @"E:\github\PlayAndStudy\src\WordTool\Doc\测试文档.docx";

            WordExport wordExport = new WordExport();
            wordExport.WordStart(temptPath);
            Console.Write("执行完成");
            Console.ReadKey();
        }
    }
}
