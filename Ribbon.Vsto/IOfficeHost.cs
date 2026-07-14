using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Ribbon.Contracts;

namespace Ribbon.Vsto
{
    public interface IOfficeHost
    {
        HostRegistration Registration { get; }
        IList<OfficeToolDefinition> GetTools();
        Task<OfficeToolResult> InvokeAsync(OfficeToolInvocation invocation, CancellationToken cancellationToken);
    }
}
