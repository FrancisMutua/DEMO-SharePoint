using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DEMO_SharePoint.Services.Interfaces;
using DEMO_SharePoint.Services.Models;

namespace DEMO_SharePoint.Services.Implementations
{
    /// <summary>
    /// Primary service implementing INetworkPrinterService
    /// Orchestrates discovery, communication, and job management services
    /// Provides unified interface for all scanner operations
    /// </summary>
    public class NetworkPrinterService : INetworkPrinterService
    {
        private readonly IScannerDiscoveryService _discoveryService;
        private readonly IScannerCommunicationService _communicationService;
        private readonly IScanJobManagementService _jobManagementService;
        private readonly List<NetworkScanner> _cachedScanners;

        public NetworkPrinterService()
            : this(null, null, null)
        {
        }

        public NetworkPrinterService(
            IScannerDiscoveryService discoveryService = null,
            IScannerCommunicationService communicationService = null,
            IScanJobManagementService jobManagementService = null)
        {
            _communicationService = communicationService ?? new ScannerCommunicationService();
            _discoveryService = discoveryService ?? new ScannerDiscoveryService(_communicationService);
            _jobManagementService = jobManagementService ?? new ScanJobManagementService();
            _cachedScanners = new List<NetworkScanner>();
        }

        /// <summary>
        /// Discovers all available network scanners on the corporate network
        /// </summary>
        public async Task<List<NetworkScanner>> DiscoverNetworkScannersAsync()
        {
            try
            {
                var scanners = await _discoveryService.DiscoverScannersAsync();
                _cachedScanners.Clear();
                _cachedScanners.AddRange(scanners);
                _jobManagementService.SetAvailableScanners(scanners);
                return scanners;
            }
            catch (Exception ex)
            {
                throw new Exception($"Scanner discovery failed: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Gets detailed device information for a specific scanner
        /// </summary>
        public async Task<NetworkScanner> GetScannerInfoAsync(string scannerId)
        {
            var scanner = _cachedScanners.FirstOrDefault(s => s.ScannerId == scannerId);
            if (scanner != null)
            {
                scanner.IsOnline = await _communicationService.TestScannerConnectivityAsync(scannerId);
                scanner.LastStatusCheck = DateTime.Now;
            }

            return await Task.FromResult(scanner);
        }

        /// <summary>
        /// Initiates a scan job on the specified network scanner
        /// </summary>
        public async Task<string> InitiateScanAsync(string scannerId, ScanParameters scanParams)
        {
            return await _jobManagementService.InitiateScanAsync(scannerId, scanParams);
        }

        /// <summary>
        /// Retrieves scan job status and page count
        /// </summary>
        public async Task<ScanJobStatus> GetScanStatusAsync(string jobId)
        {
            return await _jobManagementService.GetScanStatusAsync(jobId);
        }

        /// <summary>
        /// Downloads scanned pages as TIFF or PDF
        /// </summary>
        public async Task<byte[]> DownloadScanAsync(string jobId)
        {
            return await _jobManagementService.DownloadScanAsync(jobId);
        }

        /// <summary>
        /// Cancels an active scan job
        /// </summary>
        public async Task<bool> CancelScanAsync(string jobId)
        {
            return await _jobManagementService.CancelScanAsync(jobId);
        }

        /// <summary>
        /// Tests connectivity to scanner (ping)
        /// </summary>
        public async Task<bool> TestScannerConnectivityAsync(string scannerId)
        {
            return await _communicationService.TestScannerConnectivityAsync(scannerId);
        }

        /// <summary>
        /// Gets the cached scanner list
        /// </summary>
        public List<NetworkScanner> GetCachedScanners()
        {
            return _cachedScanners;
        }
    }
}