using System.Collections.Generic;
using System.Threading.Tasks;
using DEMO_SharePoint.Services.Models;

namespace DEMO_SharePoint.Services.Interfaces
{
    /// <summary>
    /// Service for scanner device communication
    /// Encapsulates device info retrieval and connectivity testing
    /// </summary>
    public interface IScannerCommunicationService
    {
        Task<NetworkScanner> GetDeviceInfoAsync(string ipAddress);
        Task<bool> TestScannerConnectivityAsync(string scannerId);
    }
}