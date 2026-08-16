using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Ribbon.Contracts;

namespace Ribbon.Vsto
{
    /// <summary>
    /// Core host surface every Office slice implements. Hosts are either document
    /// hosts (Excel, Word, PowerPoint) or item hosts (Outlook); only the shape of
    /// their context anchor and capabilities differ.
    /// </summary>
    public interface IOfficeHost
    {
        HostRegistration Registration { get; }
        IList<OfficeToolDefinition> GetTools();
        Task<OfficeToolResult> InvokeAsync(OfficeToolInvocation invocation, CancellationToken cancellationToken);
    }

    /// <summary>
    /// Optional capability for hosts whose working context can be snapshotted and
    /// restored in place. Item hosts such as Outlook must not implement this.
    /// </summary>
    public interface ICheckpointHost : IOfficeHost
    {
        Task<DocumentCheckpoint> CreateCheckpointAsync(string label, CancellationToken cancellationToken);
        Task RestoreCheckpointAsync(DocumentCheckpoint checkpoint, CancellationToken cancellationToken);
    }
}
