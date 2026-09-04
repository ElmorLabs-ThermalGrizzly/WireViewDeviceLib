namespace WireView2.Device
{
    public class WireView2PhanteksDevice : WireView2Device
    {
        public WireView2PhanteksDevice(string portName, int baud = 115200) : base(portName, baud, "Thermal Grizzly WireView II", "WireView II Phanteks Edition", 0xEF, 0x08) { }

    }
}
