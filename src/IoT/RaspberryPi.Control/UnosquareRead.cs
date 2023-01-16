using Iot.Device.Bmxx80;
using System;
using System.Device.Gpio;
using System.Device.I2c;
using System.Threading;
using Iot.Device.Bmxx80.PowerMode;
using System.Threading.Tasks;
using Unosquare.RaspberryIO;
using Unosquare.RaspberryIO.Abstractions;
using Unosquare.WiringPi;

namespace RaspberryPi.Control
{
    public class UnosquareRead
    {

        public static void Init()
        {
            Pi.Init<BootstrapWiringPi>();
        }

        public static void TestLedBlinking()
        {
            Init();

            var blinkingPin = Pi.Gpio[BcmPin.Gpio17];

            // Configure the pin as an output
            blinkingPin.PinMode = GpioPinDriveMode.Output;

            // perform writes to the pin by toggling the isOn variable
            var isOn = false;
            for (var i = 0; i < 20; i++)
            {
                Console.WriteLine("isOn:{0}",isOn);

                isOn = !isOn;

                blinkingPin.Write(isOn);

                Console.WriteLine("isOn:{0}",isOn);

                System.Threading.Thread.Sleep(500);
                
                Console.WriteLine("Thread Sleep End");
            }
        }
    }
}

