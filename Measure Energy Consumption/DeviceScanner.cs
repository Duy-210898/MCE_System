using System;
using System.Collections.Generic;
using System.Net.Sockets;
using System.Threading.Tasks;
using System.Linq;

namespace Measure_Energy_Consumption
{
    public class DeviceScanner
    {
        public async Task<List<string>> ScanDeltaDevices(params string[] baseIPs)
        {
            List<string> deltaDeviceIPs = new List<string>();

            var tasks = new List<Task>();

            foreach (var baseIP in baseIPs)
            {
                for (int i = 1; i <= 255; i++)
                {
                    string ipAddress = $"{baseIP}.{i}";

                    tasks.Add(Task.Run(async () =>
                    {
                        if (await IsDeltaDevice(ipAddress))
                        {
                            lock (deltaDeviceIPs)
                            {
                                deltaDeviceIPs.Add(ipAddress);
                            }
                        }
                    }));
                }
            }

            await Task.WhenAll(tasks);

            return deltaDeviceIPs;
        }

        private async Task<bool> IsDeltaDevice(string ipAddress)
        {
            try
            {
                using (TcpClient client = new TcpClient())
                {
                    Task connectTask = client.ConnectAsync(ipAddress, 12346);

                    if (await Task.WhenAny(connectTask, Task.Delay(500)) == connectTask && client.Connected)
                    {
                        return true;
                    }
                }
            }
            catch (Exception)
            {
            }

            return false;
        }
    }
}
