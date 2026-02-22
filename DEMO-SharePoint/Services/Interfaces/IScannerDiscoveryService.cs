using System.Collections.Generic;
using System.Threading.Tasks;
using DEMO_SharePoint.Services.Models;

namespace DEMO_SharePoint.Services.Interfaces
{
    /// <summary>
    /// Service for scanner discovery operations
    /// Encapsulates WSD and SNMP discovery protocols
    /// </summary>
    public interface IScannerDiscoveryService
    {
        Task<List<NetworkScanner>> DiscoverScannersAsync();
        Task<List<NetworkScanner>> DiscoverViaWSDAsync();
        Task<List<NetworkScanner>> DiscoverViaSNMPAsync();
    }
}