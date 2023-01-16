using Iot.Device.Bmxx80;
using System;
using System.Device.Gpio;
using System.Device.I2c;
using System.Threading;
using Iot.Device.Bmxx80.PowerMode;
using System.Threading.Tasks;

namespace RaspberryPi.Control
{

    public class Nature
    {
        ///
        public static void DiodeRun()
        {
            int pin = 18;
            int lightTime = 1000;
            int dimTime = 200;

            using GpioController controller = new();
            controller.OpenPin(pin, PinMode.Output);
            Console.WriteLine($"GPIO pin enabled: {pin}");
            
            while (true)
            {
                Console.WriteLine($"Light for {lightTime}ms");
                controller.Write(pin, PinValue.High);
                Thread.Sleep(lightTime);

                Console.WriteLine($"Dim for {dimTime}ms");
                controller.Write(pin, PinValue.Low);
                Thread.Sleep(dimTime);
            }
        }


        static int[] db11Data = { 0, 0, 0, 0, 0 };
        static int pinIndex = 18;
        static GpioController controller = new(PinNumberingScheme.Board);
        public static async void DHt11Start()
        {
            try
            {
                Console.WriteLine("DHt11Start.......");

                if (!controller.IsPinOpen(pinIndex))
                {
                    controller.OpenPin(pinIndex);

                    Console.WriteLine("打开PingIndex {0}", pinIndex);
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
                if (controller.IsPinOpen(pinIndex))
                {
                    controller.ClosePin(pinIndex);

                    Console.WriteLine("关闭PingIndex {0}", pinIndex);
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

            Console.WriteLine("controller PinMode {0} PinMode.Output", PinMode.Output);

            controller.Write(pinIndex, PinValue.Low);

            Console.WriteLine("controller.Write {0} PinValue.Low", pinIndex);

            Thread.Sleep(18);

            controller.Write(pinIndex, PinValue.High);

            Console.WriteLine("controller.Write {0} PinValue.High", pinIndex);

            WaitMicroseConds(40);

            controller.SetPinMode(pinIndex, PinMode.Input);

            Console.WriteLine("controller PinMode {0} PinMode.Input", PinMode.Input);

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

                Console.WriteLine("lastState Value {0}", lastState);

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
    }
}