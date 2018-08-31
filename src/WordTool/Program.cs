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
            string temptPath = @"E:\github\PlayAndStudy\src\WordTool\Doc\t2.docx";
            try
            {
                WordExport wordExport = new WordExport();
                //wordExport.WordStart(temptPath);

                wordExport.TestExport(temptPath);
            }
            catch (Exception ex)
            {
                Console.Write(ex.ToString());
            }            
            Console.Write("执行完成");
            Console.ReadKey();
        }
    }
}
