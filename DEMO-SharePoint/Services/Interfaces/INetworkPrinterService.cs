using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DEMO_SharePoint.Services.Models;

namespace DEMO_SharePoint.Services.Interfaces
{
    /// <summary>
    /// Interface for communicating with network-connected scanners and printers
    /// Supports WSD (Web Services for Devices) protocol for scanner discovery
    /// </summary>
    public interface INetworkPrinterService
    {
        /// <summary>
        /// Discovers all available network scanners on the corporate network
        /// Uses WSD protocol for device discovery
        /// </summary>
        Task<List<NetworkScanner>> DiscoverNetworkScannersAsync();

        /// <summary>
        /// Gets detailed device information for a specific scanner
        /// </summary>
        Task<NetworkScanner> GetScannerInfoAsync(string scannerId);

        /// <summary>
        /// Initiates a scan job on the specified network scanner
        /// Returns job ID for tracking
        /// </summary>
        Task<string> InitiateScanAsync(string scannerId, ScanParameters scanParams);

        /// <summary>
        /// Retrieves scan job status and page count
        /// </summary>
        Task<ScanJobStatus> GetScanStatusAsync(string jobId);

        /// <summary>
        /// Downloads scanned pages as TIFF or PDF
        /// </summary>
        Task<byte[]> DownloadScanAsync(string jobId);

        /// <summary>
        /// Cancels an active scan job
        /// </summary>
        Task<bool> CancelScanAsync(string jobId);

        /// <summary>
        /// Tests connectivity to scanner (ping)
        /// </summary>
        Task<bool> TestScannerConnectivityAsync(string scannerId);
    }
}