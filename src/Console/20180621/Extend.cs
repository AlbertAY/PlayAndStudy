using AutoMapper;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
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
