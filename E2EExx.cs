using System;
using System.IO.Ports;
using System.Threading;
using System.Collections.Generic;
using System.Text;

namespace Bev.Instruments.EplusE.E2EExx
{
    public class E2EExx
    {
        private static SerialPort comPort;
        private const string genericString = "???";     // returned if something failed
        private int delayTimeForRespond = 400;    // rather long delay necessary
        // https://docs.microsoft.com/en-us/dotnet/api/system.io.ports.serialport.close?view=dotnet-plat-ext-5.0
        private const int waitOnClose = 100;             // No actual value is given, experimental


        public E2EExx(string portName)
        {
            DevicePort = portName.Trim();
            comPort = new SerialPort(DevicePort, 9600);
            comPort.RtsEnable = true;   // this is essential
            comPort.DtrEnable = true;	// this is essential
            OpenPort();
        }

        public string DevicePort { get; }
        public string InstrumentManufacturer => "E+E Elektronik";
        public string InstrumentType => GetInstrumentType();
        public string InstrumentSerialNumber => GetInstrumentSerialNumber();
        public string InstrumentFirmwareVersion => GetInstrumentVersion();
        public string InstrumentID => $"{InstrumentType} {InstrumentFirmwareVersion} SN:{InstrumentSerialNumber} @ {DevicePort}";

        private double Temperature { get; set; }
        private double Humidity { get; set; }

        public int DelayTimeForRespond { get => delayTimeForRespond; set => delayTimeForRespond = value; }

        public MeasurementValues GetValues()
        {
            UpdateValues();
            return new MeasurementValues(Temperature, Humidity);
        }

        private void ClearCachedValues()
        {
            Temperature = double.NaN;
            Humidity = double.NaN;
        }

        private void UpdateValues()
        {
            ClearCachedValues();
            byte humLowByte = QueryE2(0x81);
            byte humHighByte = QueryE2(0x91);
            byte tempLowByte = QueryE2(0xA1);
            byte tempHighByte = QueryE2(0xB1);
            byte statusByte = QueryE2(0x71);
            if (statusByte != 0x00)
            {
                Console.WriteLine($">>> {statusByte,2:X2}");
                return;
            }
                
            Humidity = ((uint)humLowByte + (uint)humHighByte * 256) / 100.0;
            Temperature = ((uint)tempLowByte + (uint)tempHighByte * 256) / 100.0 - 273.15;
        }

        private string GetInstrumentType()
        {
            byte groupLowByte = QueryE2(0x11);
            byte subGroupByte = QueryE2(0x21);
            byte groupHighByte = QueryE2(0x41);

            if (groupHighByte == 0x55 || groupHighByte == 0xFF)
                groupHighByte = 0x00;
            int productSeries = groupHighByte * 256 + groupLowByte;
            int outputType = (subGroupByte >> 4) & 0x0F;
            int ftType = subGroupByte & 0x0F;
            string typeAsString = "EE";
            if (productSeries >= 100)
                typeAsString += $"{productSeries}";
            else
                typeAsString += $"{productSeries:00}";
            if (outputType != 0)
                typeAsString += $"-{outputType}";
            typeAsString += $" FT{ftType}";
            return typeAsString;
        }

        private string GetInstrumentSerialNumber()
        {
            return genericString;
        }

        private string GetInstrumentVersion()
        {
            return genericString;
        }

        private byte QueryE2(byte address)
        {
            //OpenPort(); //Thread.Sleep(delayTimeForRespond*2);
            SendSerialBus(ComposeCommand(address));
            Thread.Sleep(delayTimeForRespond);
            byte response = ReadByte();
            //ClosePort(); //Thread.Sleep(delayTimeForRespond);
            return response;
        }

        private byte[] ComposeCommand(byte address)
        {
            byte[] buffer = new byte[4];
            buffer[0] = 0x51;       //[B]
            buffer[1] = 0x01;       //[L]
            buffer[2] = address;    //[D]
            buffer[3] = (byte)(buffer[0] + buffer[1] + buffer[2]); //[C]
            return buffer;
        }

        private void SendSerialBus(byte[] command)
        {
            try
            {
                comPort.Write(command, 0, command.Length);
                return;
            }
            catch (Exception)
            {
                return;
            }
        }

        private byte ReadByte()
        {
            byte errorByte = 0xFF;
            try
            {
                byte[] buffer = new byte[comPort.BytesToRead];
                comPort.Read(buffer, 0, buffer.Length);
                Console.WriteLine($">>> ReadByte -> {BytesToString(buffer)}");
                if (buffer.Length != 6)
                    return errorByte;
                // TODO check syntax of response
                return buffer[4];
            }
            catch (Exception)
            {
                return errorByte;
            }
        }

        private void OpenPort()
        {
            try
            {
                if (!comPort.IsOpen)
                    comPort.Open();
            }
            catch (Exception)
            { }
        }

        private void ClosePort()
        {
            try
            {
                if (comPort.IsOpen)
                {
                    comPort.Close();
                    Thread.Sleep(waitOnClose);
                }
            }
            catch (Exception)
            { }
        }

        // function for debbuging purposes
        private string BytesToString(byte[] bytes)
        {
            string str = "";
            foreach (byte b in bytes)
                str += $" {b,2:X2}";
            return str;
        }

    }
}
