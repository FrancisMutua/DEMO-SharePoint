using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net.Sockets;
using System.Threading;
using System.Threading.Tasks;

namespace DEMO_SharePoint.Services.Base
{
    /// <summary>
    /// Abstract base class for network device communication
    /// Provides common functionality for scanner discovery and device management
    /// </summary>
    public abstract class NetworkDeviceServiceBase
    {
        protected readonly string NetworkSubnet;
        protected readonly int DevicePort;
        protected readonly int DiscoveryTimeout;

        protected NetworkDeviceServiceBase()
        {
            NetworkSubnet = ConfigurationManager.AppSettings["ScannerNetworkSubnet"] ?? "192.168.1.0/24";
            DevicePort = int.Parse(ConfigurationManager.AppSettings["ScannerPort"] ?? "9100");
            DiscoveryTimeout = int.Parse(ConfigurationManager.AppSettings["ScannerDiscoveryTimeout"] ?? "30");
        }

        /// <summary>
        /// Parses CIDR notation network subnet to individual IP addresses
        /// </summary>
        protected List<string> ParseNetworkSubnet(string subnet)
        {
            var ips = new List<string>();

            try
            {
                var parts = subnet.Split('/');
                if (parts.Length != 2)
                    return ips;

                var baseIp = parts[0];
                var cidr = int.Parse(parts[1]);

                var ipParts = baseIp.Split('.');
                if (ipParts.Length != 4)
                    return ips;

                var baseNum = (int.Parse(ipParts[0]) << 24) +
                              (int.Parse(ipParts[1]) << 16) +
                              (int.Parse(ipParts[2]) << 8) +
                              int.Parse(ipParts[3]);

                var mask = (int)(0xFFFFFFFFU << (32 - cidr));
                var networkNum = baseNum & mask;
                var broadcast = networkNum | (~mask);

                for (int i = networkNum + 1; i < broadcast; i++)
                {
                    ips.Add($"{(i >> 24) & 0xFF}.{(i >> 16) & 0xFF}.{(i >> 8) & 0xFF}.{i & 0xFF}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Subnet parse error: {ex.Message}");
            }

            return ips;
        }

        /// <summary>
        /// Tests device connectivity using TCP with proper async/await and cancellation handling
        /// </summary>
        protected async Task<bool> IsDeviceReachableAsync(string ipAddress, int port)
        {
            TcpClient client = null;
            try
            {
                client = new TcpClient();

                using (var cts = new CancellationTokenSource(DiscoveryTimeout * 1000))
                {
                    var connectTask = client.ConnectAsync(ipAddress, port);
                    var delayTask = Task.Delay(Timeout.Infinite, cts.Token);

                    var completedTask = await Task.WhenAny(connectTask, delayTask);

                    if (completedTask == delayTask)
                    {
                        return false;
                    }

                    if (connectTask.IsFaulted)
                    {
                        return false;
                    }

                    cts.Cancel();
                    return client.Connected;
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                client?.Dispose();
            }
        }
    }
}