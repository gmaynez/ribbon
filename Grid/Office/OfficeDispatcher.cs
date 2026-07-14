using System;
using System.Threading;
using System.Threading.Tasks;

namespace Grid.Office
{
    internal sealed class OfficeDispatcher
    {
        private readonly SynchronizationContext _synchronizationContext;
        private readonly int _threadId;

        public OfficeDispatcher(SynchronizationContext synchronizationContext)
        {
            if (synchronizationContext == null)
            {
                throw new ArgumentNullException(nameof(synchronizationContext));
            }

            _synchronizationContext = synchronizationContext;
            _threadId = Thread.CurrentThread.ManagedThreadId;
        }

        public Task RunAsync(Action action, CancellationToken cancellationToken)
        {
            return RunAsync<object>(delegate
            {
                action();
                return null;
            }, cancellationToken);
        }

        public Task<T> RunAsync<T>(Func<T> action, CancellationToken cancellationToken)
        {
            TaskCompletionSource<T> completionSource;

            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            cancellationToken.ThrowIfCancellationRequested();

            if (Thread.CurrentThread.ManagedThreadId == _threadId)
            {
                return Task.FromResult(action());
            }

            completionSource = new TaskCompletionSource<T>();
            _synchronizationContext.Post(delegate
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    completionSource.TrySetCanceled();
                    return;
                }

                try
                {
                    completionSource.TrySetResult(action());
                }
                catch (Exception ex)
                {
                    completionSource.TrySetException(ex);
                }
            }, null);

            return completionSource.Task;
        }
    }
}
