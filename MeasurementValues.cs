using System;

namespace Bev.Instruments.EplusE.E2EExx
{
    public class MeasurementValues : IComparable<MeasurementValues>
    {

        public DateTime TimeStamp { get; }
        public double Temperature { get; }
        public double Humidity { get; }

        public MeasurementValues(double temperature, double humidity)
        {
            TimeStamp = DateTime.UtcNow;
            Temperature = temperature;
            Humidity = humidity;
        }

        public int CompareTo(MeasurementValues other)
        {
            return TimeStamp.CompareTo(other.TimeStamp);
        }
    }
}