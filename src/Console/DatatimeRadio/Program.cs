using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatatimeRadio
{
    class Program
    {
        static void Main(string[] args)
        {
            string startTime = "2019-06-07 21:24:02.9456536";
            

            DateTime now = DateTime.Now;
            TimeSpan timeSpan = now - DateTime.Parse(startTime);
            int minutes = int.Parse("120");
            int agoMinutes = timeSpan.Days * 24 * 60 + timeSpan.Hours * 60 + timeSpan.Minutes;
            decimal radio= (decimal)agoMinutes / (decimal)minutes * 100;
            Console.WriteLine(radio);
            Console.ReadKey();
        }
    
    }
}
