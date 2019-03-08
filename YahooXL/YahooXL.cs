using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive.Concurrency;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using YahooFinanceApi;

#nullable enable

namespace YahooDelayedQuotesXLAddIn
{
    public static class YahooXL
    {
        private static readonly Dictionary<IObserver<dynamic>, (string symbol, string fieldName)> Observers = new Dictionary<IObserver<dynamic>, (string symbol, string fieldName)>();
        private static Dictionary<string, Security> Data = new Dictionary<string, Security>();
        private static readonly SemaphoreSlim Semaphore = new SemaphoreSlim(0);
        private static bool started = false;

        [ExcelFunction(Description = "Yahoo Delayed Quotes")]
        public static IObservable<dynamic> GetYahoo(string symbol, string fieldName)
        {
            return Observable.Create<dynamic>(observer => // executed once
            {
                try
                {
                    AddObserver(observer, symbol, fieldName);
                }
                catch (Exception e)
                {
                    observer.OnError(e);
                }
                return Disposable.Create(() => Observers.Remove(observer));
            });
        }
        private static void AddObserver(IObserver<dynamic> observer, string symbol, string fieldName)
        {
            bool found = Data.TryGetValue(symbol, out Security security);
            if (found)
            {
                if (security == null)
                {
                    observer.OnNext($"Symbol not found: \"{symbol}\".");
                    observer.OnCompleted();
                    return;
                }
                dynamic value = security[fieldName];
                if (value == null)
                {
                    observer.OnNext(Extensions.GetFieldNameOrNotFound(security, fieldName));
                    observer.OnCompleted();
                    return;
                }
                observer.OnNext(value);
            }

            Observers.Add(observer, (symbol, fieldName));

            //if (Interlocked.CompareExchange(ref started, 1, 0) == 0)
            if (!started)
                ImmediateScheduler.Instance.ScheduleAsync(RefreshLoop);

            if (!found)
                Semaphore.Release();
        }

        private async static Task RefreshLoop(IScheduler scheduler, CancellationToken ct)
        {
            started = true;
            try
            {
                while (Observers.Any())
                {
                    await Semaphore.WaitAsync(30000, ct).ConfigureAwait(false);
                    await Update(ct).ConfigureAwait(false);
                }
            }
            finally
            {
                //Interlocked.Exchange(ref started, 0);
                started = false;
            }
        }

        private static async Task Update(CancellationToken ct)
        {
            var symbolGroups = Observers
                .Select(kvp => (kvp.Value.symbol, kvp.Value.fieldName, kvp.Key))
                .GroupBy(x => x.symbol, StringComparer.InvariantCultureIgnoreCase)
                .ToList();

            var symbols = symbolGroups.Select(g => g.Key).ToList();

            Data = await new YahooQuotes(ct).GetAsync(symbols).ConfigureAwait(false);

            foreach (var group in symbolGroups)
            {
                string symbol = group.Key;
                Security? security = Data[symbol];
                foreach (var cell in group)
                {
                    var fieldName = cell.fieldName;
                    var observer = cell.Key;
                    if (security != null)
                    {
                        dynamic? value = security[fieldName];
                        if (value != null)
                        {
                            observer.OnNext(value);
                            continue;
                        }
                        observer.OnNext(Extensions.GetFieldNameOrNotFound(security, fieldName));
                    }
                    else
                        observer.OnNext($"Symbol not found: \"{symbol}\".");
                    observer.OnCompleted();
                    Observers.Remove(observer);
                }
            }
        }
    }
}
