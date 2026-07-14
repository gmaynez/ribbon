using System;
using System.Threading;
using System.Threading.Tasks;

namespace Ribbon.Vsto
{
    public sealed class OfficeDispatcher
    {
        private readonly SynchronizationContext _context;
        private readonly int _threadId;

        public OfficeDispatcher(SynchronizationContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _threadId = Thread.CurrentThread.ManagedThreadId;
        }

        public Task<T> RunAsync<T>(Func<T> action, CancellationToken cancellationToken)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            cancellationToken.ThrowIfCancellationRequested();
            if (Thread.CurrentThread.ManagedThreadId == _threadId) return Task.FromResult(action());

            var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            _context.Post(_ =>
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    completion.TrySetCanceled();
                    return;
                }
                try { completion.TrySetResult(action()); }
                catch (Exception exception) { completion.TrySetException(exception); }
            }, null);
            return completion.Task;
        }
    }
}
