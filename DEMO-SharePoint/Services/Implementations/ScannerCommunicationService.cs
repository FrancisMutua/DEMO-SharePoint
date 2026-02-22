using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using DEMO_SharePoint.Services.Base;
using DEMO_SharePoint.Services.Interfaces;
using DEMO_SharePoint.Services.Models;
using System.Net.Http;

namespace DEMO_SharePoint.Services.Implementations
{
    /// <summary>
    /// Implementation of scanner device communication
    /// Handles device information retrieval and connectivity testing
    /// </summary>
    public class ScannerCommunicationService : NetworkDeviceServiceBase, IScannerCommunicationService
    {
        private readonly List<NetworkScanner> _cachedScanners;

        public ScannerCommunicationService()
        {
            _cachedScanners = new List<NetworkScanner>();
        }

        /// <summary>
        /// Retrieves device information from a scanner via HTTP
        /// Most network devices have web server on port 80
        /// </summary>
        public async Task<NetworkScanner> GetDeviceInfoAsync(string ipAddress)
        {
            try
            {
                using (var client = new HttpClient(new HttpClientHandler()))
                {
                    client.Timeout = TimeSpan.FromSeconds(DiscoveryTimeout);
                    await client.GetAsync($"http://{ipAddress}");
                }

                var scanner = new NetworkScanner
                {
                    ScannerId = Guid.NewGuid().ToString(),
                    IPAddress = ipAddress,
                    FriendlyName = $"Scanner_{ipAddress}",
                    IsOnline = true,
                    LastStatusCheck = DateTime.Now,
                    SupportsDuplex = true,
                    SupportsADF = true
                };

                return await Task.FromResult(scanner);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Tests connectivity to a specific scanner
        /// </summary>
        public async Task<bool> TestScannerConnectivityAsync(string scannerId)
        {
            var scanner = _cachedScanners.FirstOrDefault(s => s.ScannerId == scannerId);
            if (scanner == null)
                return false;

            return await IsDeviceReachableAsync(scanner.IPAddress, 80);
        }

        /// <summary>
        /// Caches scanner information for quick lookup
        /// </summary>
        public void CacheScanners(List<NetworkScanner> scanners)
        {
            _cachedScanners.Clear();
            _cachedScanners.AddRange(scanners);
        }

        /// <summary>
        /// Gets cached scanner list
        /// </summary>
        public List<NetworkScanner> GetCachedScanners()
        {
            return _cachedScanners;
        }
    }
}