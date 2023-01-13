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

            // int pin = 18;
            // int lightTime = 1000;
            // int dimTime = 200;

            // using GpioController controller = new();
            // controller.OpenPin(pin, PinMode.Output);
            // Console.WriteLine($"GPIO pin enabled: {pin}");

            // //控制二极管
            // Task.Run(() =>
            // {
            //     while (true)
            //     {
            //         Console.WriteLine($"Light for {lightTime}ms");
            //         controller.Write(pin, PinValue.High);
            //         Thread.Sleep(lightTime);

            //         Console.WriteLine($"Dim for {dimTime}ms");
            //         controller.Write(pin, PinValue.Low);
            //         Thread.Sleep(dimTime);
            //     }

            // });

            Console.WriteLine($"Start Get Tem");

            Task.Run(() => DHt11Start());


            Console.ReadLine();

            return 0;

        }

        static int[] db11Data = { 0, 0, 0, 0, 0 };
        static int pinIndex = 18;
        static GpioController controller = new(PinNumberingScheme.Board);
        public static async void DHt11Start()
        {
            try
            {
                Console.WriteLine("DHt11Start.......");

                if(!controller.IsPinOpen(pinIndex))
                {
                    controller.OpenPin(pinIndex);

                    Console.WriteLine("打开PingIndex {0}",pinIndex);
                }                

                while (ReadData(controller, pinIndex))
                {
                    await Task.Delay(1000);

                    if (ReadData(controller, pinIndex))
                    {
                        Console.WriteLine("Read Data Success");

                        Console.WriteLine($"湿度：{db11Data[0]}.{db11Data[0]}%   温度");
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"DHt11Start Exception  Message{0} , Content : {1}", ex.Message, ex.StackTrace);
            }
            finally
            {
                if(controller.IsPinOpen(pinIndex))
                {
                    controller.ClosePin(pinIndex);

                    Console.WriteLine("关闭PingIndex {0}",pinIndex);
                }

                Console.WriteLine("finally");
            }
        }


        public static bool ReadData(GpioController controller, int pinIndex)
        {
            Console.WriteLine($"Read Data ........");

            PinValue lastState = PinValue.High;

            int i, j = 0;

            db11Data[0] = db11Data[1] = db11Data[2] = db11Data[3] = db11Data[4] = 0;

            controller.SetPinMode(pinIndex, PinMode.Output);

            Console.WriteLine("controller PinMode {0} PinMode.Output",PinMode.Output);

            controller.Write(pinIndex, PinValue.Low);

            Console.WriteLine("controller.Write {0} PinValue.Low",pinIndex);

            Thread.Sleep(18);

            controller.Write(pinIndex, PinValue.High);

            Console.WriteLine("controller.Write {0} PinValue.High",pinIndex);

            WaitMicroseConds(40);

            controller.SetPinMode(pinIndex, PinMode.Input);

            Console.WriteLine("controller PinMode {0} PinMode.Input",PinMode.Input);

            for (i = 0; i < 85; i++)
            {
                int conter = 0;

                while (controller.Read(pinIndex) == lastState)
                {
                    conter++;

                    WaitMicroseConds(1);

                    if (conter == 255)
                    {
                        break;
                    }
                }

                lastState = controller.Read(pinIndex);

                Console.WriteLine("lastState Value {0}",lastState);

                if (conter == 255)
                {
                    break;
                }

                if ((i >= 4) && i % 2 == 0)
                {
                    db11Data[j / 8] <<= 1;

                    if (conter > 16)
                    {
                        db11Data[j / 8] |= 1;
                    }

                    j++;
                }

            }

            Console.WriteLine($"End Read ........");

            if ((j >= 40) && (db11Data[4] == ((db11Data[0] + db11Data[1] + db11Data[2] + db11Data[3]) & 0xFF)))
            {
                return true;
            }
            else
            {
                return false;
            }

        }



        private static void WaitMicroseConds(int microseconds)
        {
            Console.WriteLine($"Start Wait MicroseConds: {microseconds} ........");

            var until = DateTime.UtcNow.Ticks + (microseconds * 10);

            while (DateTime.UtcNow.Ticks < until)
            {

            }

            Console.WriteLine($"End Wait MicroseConds: {microseconds} ........");
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





