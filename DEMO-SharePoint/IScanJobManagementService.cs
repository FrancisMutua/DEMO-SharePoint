using System.Threading.Tasks;
using DEMO_SharePoint.Services.Models;

namespace DEMO_SharePoint.Services.Interfaces
{
    /// <summary>
    /// Service for managing scan jobs
    /// Encapsulates job lifecycle: initiate, status, download, cancel
    /// </summary>
    public interface IScanJobManagementService
    {
        Task<string> InitiateScanAsync(string scannerId, ScanParameters scanParams);
        Task<ScanJobStatus> GetScanStatusAsync(string jobId);
        Task<byte[]> DownloadScanAsync(string jobId);
        Task<bool> CancelScanAsync(string jobId);
    }
}