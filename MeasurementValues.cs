using System;

namespace Bev.Instruments.EplusE.E2EExx
{
    public class MeasurementValues : IComparable<MeasurementValues>
    {

        public DateTime TimeStamp { get; }
        public double Temperature { get; }
        public double Humidity { get; }
        public double Value3 { get; }
        public double Value4 { get; }

        public MeasurementValues(double temperature, double humidity, double value3, double value4)
        {
            TimeStamp = DateTime.UtcNow;
            Temperature = temperature;
            Humidity = humidity;
            Value3 = value3;
            Value4 = value4;
        }

        public int CompareTo(MeasurementValues other)
        {
            return TimeStamp.CompareTo(other.TimeStamp);
        }
    }
}