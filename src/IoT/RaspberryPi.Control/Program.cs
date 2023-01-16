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
            Console.WriteLine($"Start The Programe!");

            Task.Run(() => UnosquareRead.TestLedBlinking());


            Console.WriteLine($"End The Programe!");    

            Console.ReadLine();

            return 0;

        }

        

        public static void Bme280Start()
        {

            try
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
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error Message: {0}, DetailInfo : {2}", ex.Message, ex.StackTrace);
            }

        }
    }
}





