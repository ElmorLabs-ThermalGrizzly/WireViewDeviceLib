using System;
using System.IO.Ports;
using System.Numerics;
using System.Threading;

namespace WireView2.Device
{
    /// <summary>
    /// A <see cref="SerialPort"/> wrapper that serializes access to USB sensor devices
    /// across processes using a named, global mutex.
    /// </summary>
    internal sealed class SharedSerialPort : SerialPort
    {
        private const string MutexName = @"Global\Access_USB_Sensors";
        private readonly Mutex _mutex = new Mutex(false, MutexName);

        /// <summary>
        /// Default wait time when acquiring the mutex. Override via ctor if needed.
        /// </summary>
        private int MutexTimeout { get; set; } = 2000; // ms
        private bool hasMutex = false;

        public SharedSerialPort()
        {
        }

        public SharedSerialPort(string portName) : base(portName)
        {
        }

        public SharedSerialPort(string portName, int baudRate) : base(portName, baudRate)
        {
        }

        public SharedSerialPort(string portName, int baudRate, Parity parity) : base(portName, baudRate, parity)
        {
        }

        public SharedSerialPort(string portName, int baudRate, Parity parity, int dataBits)
            : base(portName, baudRate, parity, dataBits)
        {
        }

        public SharedSerialPort(string portName, int baudRate, Parity parity, int dataBits, StopBits stopBits)
            : base(portName, baudRate, parity, dataBits, stopBits)
        {
        }

        public new void Open()
        {
            try
            {
                _mutex.WaitOne(MutexTimeout);
                hasMutex = true;
            }
            catch (AbandonedMutexException)
            {
                // Another process terminated without releasing the mutex.
                // We can still acquire it, so just proceed.
                //base.Open();
                _mutex.ReleaseMutex();
            }
            if(hasMutex) base.Open();
        }

        public new void Close()
        {
            if (hasMutex)
            {

                BaseStream.Flush();
                BaseStream.Close();
                try
                {
                    hasMutex = false;
                    _mutex.ReleaseMutex();
                }
                catch { }
            }
        }

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing)
                {
                    // Don’t guard Dispose with the mutex; Dispose may be called during teardown and
                    // we want to avoid deadlocks if another thread/process is misbehaving.
                    _mutex.Dispose();
                }
            }
            finally
            {
                base.Dispose(disposing);
            }
        }

    }
}
