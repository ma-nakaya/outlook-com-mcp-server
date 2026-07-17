using System.Collections.Concurrent;

namespace OutlookComMcp.Outlook;

public sealed class StaDispatcher : IDisposable
{
    private readonly BlockingCollection<IWorkItem> _queue = [];
    private readonly Thread _thread;
    private bool _disposed;

    public StaDispatcher()
    {
        _thread = new Thread(Run)
        {
            IsBackground = true,
            Name = "Outlook COM STA",
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    public Task<T> InvokeAsync<T>(Func<T> action, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentNullException.ThrowIfNull(action);

        if (cancellationToken.IsCancellationRequested)
        {
            return Task.FromCanceled<T>(cancellationToken);
        }

        WorkItem<T> workItem = new(action, cancellationToken);
        _queue.Add(workItem, cancellationToken);
        return workItem.Task;
    }

    private void Run()
    {
        foreach (IWorkItem workItem in _queue.GetConsumingEnumerable())
        {
            workItem.Execute();
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _queue.CompleteAdding();
        _thread.Join(TimeSpan.FromSeconds(5));
        _queue.Dispose();
    }

    private interface IWorkItem
    {
        void Execute();
    }

    private sealed class WorkItem<T>(Func<T> action, CancellationToken cancellationToken) : IWorkItem
    {
        private readonly TaskCompletionSource<T> _completion =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        public Task<T> Task => _completion.Task;

        public void Execute()
        {
            if (cancellationToken.IsCancellationRequested)
            {
                _completion.TrySetCanceled(cancellationToken);
                return;
            }

            try
            {
                _completion.TrySetResult(action());
            }
            catch (Exception exception)
            {
                _completion.TrySetException(exception);
            }
        }
    }
}

