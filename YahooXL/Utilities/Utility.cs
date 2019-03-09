using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using YahooFinanceApi;

#nullable enable

namespace YahooXL
{
    internal static class Extensions
    {
        internal static string GetFieldNameOrNotFound(Security security, string fieldName) =>
           int.TryParse(fieldName, out int i) ? GetFieldNameFromIndex(security, i) : $"Field not found: \"{fieldName}\".";

        private static string GetFieldNameFromIndex(Security security, int index) => // slow!
            security.Fields.Keys.OrderBy(k => k).ElementAtOrDefault(index) ?? $"Invalid field index: {index}.";
    }

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
            if (!await Semaphore.WaitAsync(timeout, cancellationToken))
                return (false, new T());

            Queue.TryDequeue(out T item);
            return (true, item);
        }

        public void Dispose() => Semaphore.Dispose();
    }

}
