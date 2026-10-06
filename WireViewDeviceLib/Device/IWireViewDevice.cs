namespace WireView2.Device
{
    public interface IWireViewDevice
    {
        bool Connected { get; }
        string DeviceName { get; }
        string HardwareRevision { get; }
        string FirmwareVersion { get; }
        string UniqueId { get; }
        PowerControlType PowerControl { get; }

        void Connect();
        void Disconnect();
        event EventHandler<DeviceData>? DataUpdated;
        event EventHandler<bool>? ConnectionChanged;
    }

    // How the device acts on the system when a fault is configured to power down
    public enum PowerControlType
    {
        None,        // No power control
        PowerButton, // Short/long press of the system power button (soft/hard power)
        GpuSenseCut  // Disconnects the 12V-2x6 sense lines, GPU loses power (hard power only)
    }

    public class DeviceData
    {
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
        public bool Connected { get; set; }
        public string HardwareRevision { get; set; } = "A0";
        public string FirmwareVersion { get; set; } = "0.0.0";

        public double[] PinVoltage { get; set; } = new double[6];
        public double[] PinCurrent { get; set; } = new double[6];

        public double OnboardTempInC { get; set; }
        public double OnboardTempOutC { get; set; }
        public double ExternalTemp1C { get; set; }
        public double ExternalTemp2C { get; set; }

        public int PsuCapabilityW { get; set; }

        public double SumCurrentA => PinCurrent.Sum();
        public double SumPowerW => PinVoltage.Zip(PinCurrent, (v, i) => v * i).Sum();

        public ushort FaultStatus { get; set; }
        public ushort FaultLog { get; set; }
    }
}
