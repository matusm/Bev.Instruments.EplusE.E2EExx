using System;
using System.IO.Ports;
using System.Threading;

namespace Bev.Instruments.EplusE.E2EExx
{
    public class E2EExx
    {
        private static SerialPort comPort;
        private const string defaultString = "???"; // returned if something failed
        private int delayTimeForRespond = 400;      // rather long delay necessary
        private const int delayOnPortClose = 100;   // No actual value is given, experimental
        private const int delayOnPortOpen = 100;

        public E2EExx(string portName)
        {
            DevicePort = portName.Trim();
            comPort = new SerialPort(DevicePort, 9600);
            comPort.RtsEnable = true;   // this is essential
            comPort.DtrEnable = true;	// this is essential
        }

        public string DevicePort { get; }
        public string InstrumentManufacturer => "E+E Elektronik";
        public string InstrumentType => GetInstrumentType();
        public string InstrumentSerialNumber => GetInstrumentSerialNumber();
        public string InstrumentFirmwareVersion => GetInstrumentVersion();
        public string InstrumentID => $"{InstrumentType} {InstrumentFirmwareVersion} SN:{InstrumentSerialNumber} @ {DevicePort}";

        private double Temperature { get; set; }
        private double Humidity { get; set; }
        public double Value3 { get; set; }
        public double Value4 { get; set; }

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
            byte? humLowByte = QueryE2(0x81);
            byte? humHighByte = QueryE2(0x91);
            byte? tempLowByte = QueryE2(0xA1);
            byte? tempHighByte = QueryE2(0xB1);
            byte? value3LowByte = QueryE2(0xC1);
            byte? value3HighByte = QueryE2(0xD1);
            byte? value4LowByte = QueryE2(0xE1);
            byte? value4HighByte = QueryE2(0xD1);
            byte? statusByte = QueryE2(0x71);
            if (statusByte != 0x00)
                return;
            if (humLowByte.HasValue && humHighByte.HasValue)
                Humidity = (humLowByte.Value + humHighByte.Value * 256.0) / 100.0;
            if (tempLowByte.HasValue && tempHighByte.HasValue)
                Temperature = (tempLowByte.Value + tempHighByte.Value * 256.0) / 100.0 - 273.15;
            if (value3LowByte.HasValue && value3HighByte.HasValue)
                Value3 = value3LowByte.Value + value3HighByte.Value * 256.0;
            if (value4LowByte.HasValue && value4HighByte.HasValue)
                Value4 = value4LowByte.Value + value4HighByte.Value * 256.0;
        }

        private string GetInstrumentType()
        {
            byte? groupLowByte = QueryE2(0x11);
            if (!groupLowByte.HasValue)
                return defaultString;
            byte? subGroupByte = QueryE2(0x21);
            if (!subGroupByte.HasValue)
                return defaultString;
            byte? groupHighByte = QueryE2(0x41);
            if (!groupHighByte.HasValue)
                return defaultString;
            if (groupHighByte == 0x55 || groupHighByte == 0xFF)
                groupHighByte = 0x00;
            int productSeries = groupHighByte.Value * 256 + groupLowByte.Value;
            int outputType = (subGroupByte.Value >> 4) & 0x0F;
            int ftType = subGroupByte.Value & 0x0F;
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
            return defaultString;
        }

        private string GetInstrumentVersion()
        {
            return defaultString;
        }

        private byte? QueryE2(byte address)
        {
            OpenPort();
            SendSerialBus(ComposeCommand(address));
            Thread.Sleep(delayTimeForRespond);
            var response = ReadByte();
            // ClosePort();
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

        private byte? ReadByte()
        {
            try
            {
                byte[] buffer = new byte[comPort.BytesToRead];
                comPort.Read(buffer, 0, buffer.Length);
                Console.WriteLine($">>> ReadByte -> {BytesToString(buffer)}");
                if (IsIncorrect(buffer))
                    return null;
                return buffer[4];
            }
            catch (Exception)
            {
                return null;
            }
        }

        private bool IsIncorrect(byte[] buffer)
        {
            if (buffer.Length != 6)
                return true;
            if (buffer[0] != 0x51)  // [B]
                return true;
            if (buffer[1] != 0x03)  // [L]
                return true;
            if (buffer[2] != 0x06)  // [S]
                return true;
            if (buffer[3] != 0x00)  // [F]
                return true;
            byte crc = (byte)(buffer[0] + buffer[1] + buffer[2] + buffer[3] + buffer[4]);
            if (buffer[5] != crc)   // [C]
                return true;
            return false;
        }

        private void OpenPort()
        {
            try
            {
                if (!comPort.IsOpen)
                {
                    comPort.Open();
                    Thread.Sleep(delayOnPortOpen);
                }

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
                    Thread.Sleep(delayOnPortClose);
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

    public struct E2Return
    {
        public byte by;
        public bool st; // true -> error
    }

    public enum E2ErrorType
    {
        Unknown,
        NoError,
        CrcError,
        LengthError,
        NakError,
        CodeError
    }

}
