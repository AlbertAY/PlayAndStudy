using Iot.Device.Bmxx80;
using System;
using System.Device.Gpio;
using System.Device.I2c;
using System.Threading;
using Iot.Device.Bmxx80.PowerMode;
using System.Threading.Tasks;
using Unosquare.RaspberryIO;
using Unosquare.RaspberryIO.Abstractions;

namespace RaspberryPi.Control
{
    public class UnosquareRead
    {
        public static void TestLedBlinking()
        {
            // Get a reference to the pin you need to use.
            // Both methods below are equivalent
            var blinkingPin = Pi.Gpio[17];
            blinkingPin = Pi.Gpio[BcmPin.Gpio17];

            // Configure the pin as an output
            blinkingPin.PinMode = GpioPinDriveMode.Output;

            // perform writes to the pin by toggling the isOn variable
            var isOn = false;
            for (var i = 0; i < 20; i++)
            {
                isOn = !isOn;
                blinkingPin.Write(isOn);
                System.Threading.Thread.Sleep(500);
            }
        }
    }
}

