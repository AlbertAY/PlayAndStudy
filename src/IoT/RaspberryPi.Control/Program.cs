using Iot.Device.Bmxx80;
using System;
using System.Device.Gpio;
using System.Device.I2c;
using System.Threading;
using Iot.Device.Bmxx80.PowerMode;
using System.Threading.Tasks;

namespace RaspberryPi.Control
{
    public class Program
    {
        public static int Main(string[] args)
        {

            int pin = 18;
            int lightTime = 1000;
            int dimTime = 200;

            using GpioController controller = new();
            controller.OpenPin(pin, PinMode.Output);
            Console.WriteLine($"GPIO pin enabled: {pin}");


            //控制二极管
            Task.Run(() =>
            {
                while (true)
                {
                    Console.WriteLine($"Light for {lightTime}ms");
                    controller.Write(pin, PinValue.High);
                    Thread.Sleep(lightTime);

                    Console.WriteLine($"Dim for {dimTime}ms");
                    controller.Write(pin, PinValue.Low);
                    Thread.Sleep(dimTime);
                }

            });

            Console.WriteLine($"Start Get Tem");

            Task.Run(()=>Bme280Start());


            Console.ReadLine();

            return 0;

        }



        public static void Bme280Start()
        {
            var busId = 1;

            var i2cSettings = new I2cConnectionSettings(busId, Bme280.DefaultI2cAddress);

            using var i2cDevice = I2cDevice.Create(i2cSettings);

            var bme280 = new Bme280(i2cDevice);

            int measurementTime = bme280.GetMeasurementDuration();


            Console.WriteLine($"AAAAAA");

            while (true)
            {                

                Console.WriteLine($"BBBBB");
                
                bme280.SetPowerMode(Bmx280PowerMode.Forced);

                Thread.Sleep(measurementTime);

                bme280.TryReadTemperature(out var tempValue);
                bme280.TryReadPressure(out var preValue);
                bme280.TryReadHumidity(out var humValue);
                bme280.TryReadAltitude(out var altValue);

                Console.WriteLine($"Temperature: {tempValue.DegreesCelsius:0.#}\u00B0C");
                Console.WriteLine($"Pressure: {preValue.Hectopascals:#.##} hPa");
                Console.WriteLine($"Relative humidity: {humValue.Percent:#.##}%");
                Console.WriteLine($"Estimated altitude: {altValue.Meters:#} m");

                Thread.Sleep(1000);

            }

            
        }
    }
}





