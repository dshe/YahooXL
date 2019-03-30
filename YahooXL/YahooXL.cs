using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Reactive.Concurrency;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using YahooFinanceApi;
using System.Collections.Concurrent;

#nullable enable

namespace YahooXL
{
    public static class YahooQuotesAddin
    {
        private static readonly IScheduler Scheduler = NewThreadScheduler.Default; // CurrentThreadScheduler.Instance;

        private static readonly Dictionary<IObserver<object>, (string symbol, string fieldName)> Observers =
            new Dictionary<IObserver<object>, (string symbol, string fieldName)>();

        private static readonly AsyncBlockingCollection<(IObserver<object> observer, string symbol, string fieldName)> ObserversToAdd =
            new AsyncBlockingCollection<(IObserver<object> observer, string symbol, string fieldName)>();

        private static readonly ConcurrentBag<IObserver<object>> ObserversToRemove = new ConcurrentBag<IObserver<object>>();

        private static Dictionary<string, Security> Data = new Dictionary<string, Security>();

        private static int started = 0;
        private static IDisposable disposable;

        [ExcelFunction(Description = "Yahoo Delayed Quotes")]
        public static IObservable<object> YahooQuote([ExcelArgument("Symbol")] string symbol, [ExcelArgument("Field")] string fieldName)
        {
            if (string.IsNullOrWhiteSpace(symbol))
                return Observable.Return("Invalid symbol.");
            if (string.IsNullOrWhiteSpace(fieldName))
                fieldName = "RegularMarketPrice";
            bool found = Data.TryGetValue(symbol, out Security security);
            object? value = null;
            if (found)
            {
                if (security == null)
                    return Observable.Return($"Symbol not found: \"{symbol}\".");
                if (!security.Fields.TryGetValue(fieldName, out value))
                    return Observable.Return(Extensions.GetFieldNameOrNotFound(security, fieldName));
            }

            return Observable.Create<object>(observer => // executed once
            {
                try
                {
                    if (value != null)
                        observer.OnNext(value);

                    ObserversToAdd.Add((observer, symbol, fieldName));

                    if (Interlocked.CompareExchange(ref started, 1, 0) == 0)
                        disposable = Scheduler.ScheduleAsync(RefreshLoop);
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.ToString());
                    observer.OnError(e);
                }
                return Disposable.Create(() => ObserversToRemove.Add(observer));
            });
        }



        private static async Task RefreshLoop(IScheduler scheduler, CancellationToken ct)
        {
            try
            {
                Debug.WriteLine("starting loop");
                while (true)
                {
                    var wait = TimeSpan.FromSeconds(30);
                    while (true)
                    {
                        var (Take, Item) = await ObserversToAdd.TryTakeAsync(wait, ct).ConfigureAwait(false);
                        if (!Take)
                            break;
                        Debug.WriteLine("adding: " + Item.symbol + ", " + Item.fieldName);
                        Observers.Add(Item.observer, (Item.symbol, Item.fieldName));
                        wait = TimeSpan.FromSeconds(1);
                    }
                    Debug.WriteLine("updating");
                    await Update(ct).ConfigureAwait(false);
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.ToString());
            }
            finally
            {
                Debug.WriteLine("ending");
                Interlocked.Exchange(ref started, 0);
            }
        }

        private static async Task Update(CancellationToken ct)
        {
            while (ObserversToRemove.TryTake(out IObserver<object> observer))
                Observers.Remove(observer);

            if (!Observers.Any())
                return;

            var symbolGroups = Observers
                .Select(kvp => (kvp.Value.symbol, kvp.Value.fieldName, kvp.Key))
                .GroupBy(x => x.symbol, StringComparer.InvariantCultureIgnoreCase)
                .ToList();

            var symbols = symbolGroups.Select(g => g.Key).ToList();

            Data = await new YahooQuotes(null, ct).GetAsync(symbols).ConfigureAwait(false);

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
                        if (security.Fields.TryGetValue(fieldName, out object? value))
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
