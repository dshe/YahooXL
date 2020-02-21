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
using YahooQuotesApi;
using System.Collections.Concurrent;
using System.Reflection;

#nullable enable

namespace YahooXL
{
    public static class YahooQuotesAddin
    {
        private static readonly YahooSnapshot YahooSnapshot = new YahooSnapshot();

        private static readonly IScheduler Scheduler = NewThreadScheduler.Default;

        private static readonly Dictionary<IObserver<object>, (string symbol, string fieldName)> Observers =
            new Dictionary<IObserver<object>, (string symbol, string fieldName)>();

        private static readonly BlockingCollection<(IObserver<object> observer, string symbol, string fieldName)> ObserversToAdd =
            new BlockingCollection<(IObserver<object> observer, string symbol, string fieldName)>();

        private static readonly ConcurrentBag<IObserver<object>> ObserversToRemove = new ConcurrentBag<IObserver<object>>();

        private static Dictionary<string, Security?> Data = new Dictionary<string, Security?>();

        private static int started = 0;

        [ExcelFunction(Description = "Yahoo Delayed Quotes")]
        public static IObservable<object> YahooQuote([ExcelArgument("Symbol")] string symbol, [ExcelArgument("Field")] string fieldName)
        {
            if (string.IsNullOrWhiteSpace(symbol))
                return Observable.Return(Assembly.GetExecutingAssembly().GetName().ToString());
            if (string.IsNullOrWhiteSpace(fieldName))
                fieldName = "RegularMarketPrice";
            bool found = Data.TryGetValue(symbol, out Security? security);
            object? value = "~";
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
                    observer.OnNext(value);

                    ObserversToAdd.Add((observer, symbol, fieldName));

                    if (Interlocked.CompareExchange(ref started, 1, 0) == 0)
                        Scheduler.ScheduleAsync(RefreshLoop);
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
            Thread.CurrentThread.IsBackground = true;
            try
            {
                Debug.WriteLine("starting loop");
                while (true)
                {
                    int wait = 30000; // 30s
                    while (true)
                    {
                        if (!ObserversToAdd.TryTake(out (IObserver<object> observer, string symbol, string fieldName) Item, wait, ct))
                            break;
                        wait = 1000; // 1s
                        Debug.WriteLine("adding: " + Item.symbol + ", " + Item.fieldName);
                        Observers.Add(Item.observer, (Item.symbol, Item.fieldName));
                    }

                    while (ObserversToRemove.TryTake(out IObserver<object> observer))
                        Observers.Remove(observer);

                    if (Observers.Any())
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
                //Interlocked.Exchange(ref started, 0);
            }
        }

        private static async Task Update(CancellationToken ct)
        {
            Debug.WriteLine("updating");

            var symbolGroups = Observers
                .Select(kvp => (kvp.Value.symbol, kvp.Value.fieldName, kvp.Key))
                .GroupBy(x => x.symbol, StringComparer.InvariantCultureIgnoreCase)
                .ToList();

            var symbols = symbolGroups.Select(g => g.Key).ToList();

            Data = await YahooSnapshot.GetAsync(symbols).ConfigureAwait(false);

            foreach (var group in symbolGroups)
            {
                string symbol = group.Key;
                Security? security = Data[symbol];
                foreach (var cell in group)
                {
                    var fieldName = cell.fieldName;
                    var observer = cell.Key;
                    if (security == null)
                        observer.OnNext($"Symbol not found: {symbol}.");
                    else if (!security.Fields.TryGetValue(fieldName, out object? value))
                        observer.OnNext(Extensions.GetFieldNameOrNotFound(security, fieldName));
                    else
                    {
                        observer.OnNext(value);
                        continue;
                    }
                    observer.OnCompleted();
                    Observers.Remove(observer);
                }
            }
        }
    }
}
