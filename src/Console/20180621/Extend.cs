using AutoMapper;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _20180621
{
    public static class Extend
    {
        public static void ExtendTest<T>(T animal)
        {
            Mapper.Initialize(x => x.CreateMap<T, Animal>());
            List<Animal> data = Mapper.Map<List<Animal>>(animal);
            string strSerializeJSON = JsonConvert.SerializeObject(data);
            Console.WriteLine(strSerializeJSON);

            string aaaa = JsonConvert.SerializeObject(animal);
            Console.WriteLine(aaaa);


            Console.ReadKey();
            

        }

        public static void ExtendTest(Animal animal)
        {
            Mapper.Initialize(x => x.CreateMap<dynamic, Animal>());
            List<Animal> data = Mapper.Map<List<Animal>>(animal);
            string strSerializeJSON = JsonConvert.SerializeObject(data);
            Console.WriteLine(strSerializeJSON);

            string aaaa = JsonConvert.SerializeObject(animal);
            Console.WriteLine(aaaa);


            Console.ReadKey();


        }

        public static void TestDynamic()
        {
            Console.WriteLine($"开始。。。。。");
            int max1 = 1000000000;

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            for (int i = 0; i < max1; i++)
            {
                DynamicSample dynamicSample = new DynamicSample();
                var addMethod = typeof(DynamicSample).GetMethod("Add");
                int re = (int)addMethod.Invoke(dynamicSample, new object[] { 1, 2 });
                i++;
            }
            stopwatch.Stop();
            Console.WriteLine($"第一个执行时间:{stopwatch.ElapsedMilliseconds}");
            stopwatch.Restart();
            for (int i = 0; i < max1; i++)
            {
                dynamic dynamicSample2 = new DynamicSample();
                int re2 = dynamicSample2.Add(1, 2);
                i++;
            }
            stopwatch.Stop();
            Console.WriteLine($"第二个执行时间:{stopwatch.ElapsedMilliseconds}");
            Console.ReadKey();
        }
       
        

    }
    public class DynamicSample
    {
        public string Name { get; set; }

        public int Add(int a, int b)
        {
            return a + b;
        }
    }

    public class Animal
    {
        public string Eye { set; get; }

        public string Mouth { set; get; }

    }

    public class Person : Animal
    {
        public string Speak { set; get; }
    }


}
