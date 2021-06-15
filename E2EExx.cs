using System;
using System.IO.Ports;
using System.Threading;

namespace Bev.Instruments.EplusE.E2EExx
{
    public class E2EExx
    {
        private static SerialPort comPort;
        private const string defaultString = "???"; // returned if something failed
        private int delayTimeForRespond = 50;       // rather long delay necessary
        private const int delayOnPortClose = 50;    // No actual value is given, experimental
        private const int delayOnPortOpen = 50;

        private bool humidityAvailable;
        private bool temperatureAvailable;
        private bool airVelocityAvailable;
        private bool co2Available;

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
        private double Value3 { get; set; }
        private double Value4 { get; set; }

        public MeasurementValues GetValues()
        {
            UpdateValues();
            return new MeasurementValues(Temperature, Humidity, Value3, Value4);
        }

        private void ClearCachedValues()
        {
            Temperature = double.NaN;
            Humidity = double.NaN;
            Value3 = double.NaN;
            Value4 = double.NaN;
        }

        private void UpdateValues()
        {
            ClearCachedValues();
            byte? humLowByte = QueryE2(0x81);
            byte? humHighByte = QueryE2(0x91);
            byte? tempLowByte = QueryE2(0xA1);
            byte? tempHighByte = QueryE2(0xB1);
            byte? statusByte = QueryE2(0x71);
            if (statusByte != 0x00)
                return;
            if (humLowByte.HasValue && humHighByte.HasValue)
                Humidity = (humLowByte.Value + humHighByte.Value * 256.0) / 100.0;
            if (tempLowByte.HasValue && tempHighByte.HasValue)
                Temperature = (tempLowByte.Value + tempHighByte.Value * 256.0) / 100.0 - 273.15;
        }

        private void UpdateAllValues()
        {
            byte? humLowByte, humHighByte;
            byte? tempLowByte, tempHighByte;
            byte? value3LowByte, value3HighByte;
            byte? value4LowByte, value4HighByte;
            ClearCachedValues();
            GetAvailableValues();
            if (humidityAvailable)
            {
                humLowByte = QueryE2(0x81);
                humHighByte = QueryE2(0x91);
                if (humLowByte.HasValue && humHighByte.HasValue)
                    Humidity = (humLowByte.Value + humHighByte.Value * 256.0) / 100.0;
            }
            if (temperatureAvailable)
            {
                tempLowByte = QueryE2(0xA1);
                tempHighByte = QueryE2(0xB1);
                if (tempLowByte.HasValue && tempHighByte.HasValue)
                    Temperature = (tempLowByte.Value + tempHighByte.Value * 256.0) / 100.0 - 273.15;
            }
            if (co2Available || airVelocityAvailable)
            {
                value3LowByte = QueryE2(0xC1);
                value3HighByte = QueryE2(0xD1);
                value4LowByte = QueryE2(0xE1);
                value4HighByte = QueryE2(0xD1);
                if (value3LowByte.HasValue && value3HighByte.HasValue)
                    Value3 = value3LowByte.Value + value3HighByte.Value * 256.0;
                if (value4LowByte.HasValue && value4HighByte.HasValue)
                    Value4 = value4LowByte.Value + value4HighByte.Value * 256.0;
            }
            byte? statusByte = QueryE2(0x71);
            if (statusByte != 0x00)
                ClearCachedValues();
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

        private void GetAvailableValues()
        {
            byte? bitPattern = QueryE2(0x31);
            if(bitPattern.HasValue)
            {
                humidityAvailable = BitIsSet(bitPattern.Value, 0);
                temperatureAvailable = BitIsSet(bitPattern.Value, 1);
                airVelocityAvailable = BitIsSet(bitPattern.Value, 2);
                co2Available = BitIsSet(bitPattern.Value, 3);
            }
        }

        private bool BitIsSet(byte bitPattern, int place)
        {
            if (place < 0)
                return false;
            if (place >= 8)
                return false;
            var b = (bitPattern >> place) & 0x01;
            if (b == 0x00)
                return false;
            return true;
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
            SendCommand(ComposeCommand(address));
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

        private void SendCommand(byte[] command)
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
                // Console.WriteLine($">>> ReadByte -> {BytesToString(buffer)}");
                if (IsFaulty(buffer))
                    return null;
                return buffer[4];
            }
            catch (Exception)
            {
                return null;
            }
        }

        private bool IsFaulty(byte[] buffer)
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

}
