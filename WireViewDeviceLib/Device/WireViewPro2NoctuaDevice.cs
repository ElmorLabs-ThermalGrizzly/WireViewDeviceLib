namespace WireView2.Device
{
    public partial class WireViewPro2NoctuaDevice : WireViewPro2Device
    {

        public WireViewPro2NoctuaDevice(string portName, int baud = 115200) : base(portName, baud, "Thermal Grizzly WireView Pro II", "WireView Pro II Noctua Edition", 0xEF, 0x06) { }

    }
}