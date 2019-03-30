using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;

#nullable enable

namespace YahooXL
{
    internal class AsyncBlockingCollection<T> : IDisposable where T:new()
    {
        private readonly ConcurrentQueue<T> Queue = new ConcurrentQueue<T>();
        private readonly SemaphoreSlim Semaphore = new SemaphoreSlim(0, int.MaxValue);

        internal void Add(T item)
        {
            Queue.Enqueue(item);
            Semaphore.Release();
        }

        internal (bool Take, T Item) TryTake(TimeSpan timeout = default, CancellationToken cancellationToken = default)
        {
            if (!Semaphore.Wait(timeout, cancellationToken))
                return (false, new T());

            Queue.TryDequeue(out T item);
            return (true, item);
        }

        internal async Task<(bool Take, T Item)> TryTakeAsync(TimeSpan timeout = default, CancellationToken cancellationToken = default)
        {
            if (!await Semaphore.WaitAsync(timeout, cancellationToken).ConfigureAwait(false))
                return (false, new T());

            Queue.TryDequeue(out T item);
            return (true, item);
        }

        public void Dispose() => Semaphore.Dispose();
    }

}
