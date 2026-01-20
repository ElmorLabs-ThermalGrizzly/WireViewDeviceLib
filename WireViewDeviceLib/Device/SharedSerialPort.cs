using System;
using System.IO.Ports;
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

        // Created once per instance; mutex is named so it synchronizes across processes.
        private readonly Mutex _mutex = new Mutex(false, MutexName);

        /// <summary>
        /// Default wait time when acquiring the mutex. Override via ctor if needed.
        /// </summary>
        public TimeSpan MutexTimeout { get; set; } = TimeSpan.FromSeconds(1);

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
            using var _ = EnterMutex();
            base.Open();
        }

        public new void Close()
        {
            using var _ = EnterMutex();
            base.Close();
        }

        public new int Read(byte[] buffer, int offset, int count)
        {
            using var _ = EnterMutex();
            return base.Read(buffer, offset, count);
        }

        public new int ReadByte()
        {
            using var _ = EnterMutex();
            return base.ReadByte();
        }

        public new string ReadExisting()
        {
            using var _ = EnterMutex();
            return base.ReadExisting();
        }

        public new string ReadLine()
        {
            using var _ = EnterMutex();
            return base.ReadLine();
        }

        public new string ReadTo(string value)
        {
            using var _ = EnterMutex();
            return base.ReadTo(value);
        }

        public new void Write(byte[] buffer, int offset, int count)
        {
            using var _ = EnterMutex();
            base.Write(buffer, offset, count);
        }

        public new void Write(string text)
        {
            using var _ = EnterMutex();
            base.Write(text);
        }

        public new void WriteLine(string text)
        {
            using var _ = EnterMutex();
            base.WriteLine(text);
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

        private Releaser EnterMutex()
        {
            try
            {
                if (!_mutex.WaitOne(MutexTimeout))
                    throw new TimeoutException($"Timed out waiting for mutex '{MutexName}' within {MutexTimeout}.");
            }
            catch (AbandonedMutexException)
            {
                // Previous owner terminated without releasing; we still acquire the mutex and can proceed.
            }

            return new Releaser(_mutex);
        }

        private readonly struct Releaser : IDisposable
        {
            private readonly Mutex? _toRelease;

            public Releaser(Mutex toRelease) => _toRelease = toRelease;

            public void Dispose()
            {
                try
                {
                    _toRelease?.ReleaseMutex();
                }
                catch (ApplicationException)
                {
                    // Ignore if not owned (shouldn't happen, but avoid crashing on unexpected call paths).
                }
            }
        }
    }
}
