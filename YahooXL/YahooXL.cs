using System.Reactive.Concurrency;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using ExcelDna.Integration;
using System.Collections.Concurrent;
using System.Reflection;

namespace YahooXL;

public static class YahooQuotesAddin
{
    private static readonly List<string> PropertyNames = SecurityPropertyNames.Get();
    private static readonly YahooQuotesData YahooQuotesData = new();
    private static readonly IScheduler Scheduler = NewThreadScheduler.Default;
    private static readonly Dictionary<IObserver<object>, (string symbol, string property)> Observers = new();
    private static readonly BlockingCollection<(IObserver<object> observer, string symbol, string property)> ObserversToAdd = new();
    private static readonly ConcurrentBag<IObserver<object>> ObserversToRemove = new();
    private static Dictionary<string, Security?> Data = new(StringComparer.OrdinalIgnoreCase);
    private static int started = 0;

    [ExcelFunction(Description = "YahooXL: Delayed Quotes")]
    public static IObservable<object> YahooQuote([ExcelArgument("Symbol")] string symbol, [ExcelArgument("Property")] string property)
    {
        if (string.IsNullOrWhiteSpace(symbol) && string.IsNullOrWhiteSpace(property))
            return Observable.Return(Assembly.GetExecutingAssembly().GetName().ToString());
        if (string.IsNullOrWhiteSpace(property))
            property = "RegularMarketPrice";
        else if (int.TryParse(property, out int i))
        {
            if (i >= 0 && i < PropertyNames.Count)
                property = PropertyNames[i];
            else
                return Observable.Return($"YahooXL: invalid property index: {i}.");
        }
        else if (!PropertyNames.Contains(property, StringComparer.OrdinalIgnoreCase))
        {
            string msg = $"YahooXL: invalid property name";
            if (!property.StartsWith(msg))
                msg += $": \"{property}\"";
            return Observable.Return(msg + ".");
        }

        if (string.IsNullOrWhiteSpace(symbol))
            return Observable.Return(property);

        object? value = "~";

        // symbol has already been requested
        if (Data.TryGetValue(symbol, out Security? security))
        {
            if (security is null)
            {
                string msg = $"YahooXL: symbol not found";
                if (!symbol.StartsWith(msg))
                    msg += $": \"{symbol}\"";
                return Observable.Return(msg + ".");
            }
            value = ExcelFormatter.Get(security, property);
        }

        return Observable.Create<object>(observer => // executed once
        {
            try
            {
                observer.OnNext(value ?? "YahooXL: null value.");

                ObserversToAdd.Add((observer, symbol, property));

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
                    if (!ObserversToAdd.TryTake(out (IObserver<object> observer, string symbol, string property) Item, wait, ct))
                        break;
                    wait = 1000; // 1s
                    Observers.Add(Item.observer, (Item.symbol, Item.property));
                    Debug.WriteLine("added observer: " + Item.symbol + ", " + Item.property);
                }

                while (ObserversToRemove.TryTake(out IObserver<object>? observer))
                {
                    Observers.Remove(observer);
                    Debug.WriteLine("removed observer");
                    Debug.Flush();
                }

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
            started = 0;
        }
    }

    private static async Task Update(CancellationToken ct)
    {
        Debug.WriteLine("updating");

        var symbolGroups = Observers
            .Select(kvp => (kvp.Value.symbol, kvp.Value.property, kvp.Key))
            .GroupBy(x => x.symbol, StringComparer.OrdinalIgnoreCase)
            .ToList();

        List<string> symbols = symbolGroups.Select(g => g.Key).ToList();

        // Reference assignment is guaranteed to be atomic.
        Data = await YahooQuotesData.GetSecuritiesAsync(symbols, ct).ConfigureAwait(false);

        foreach (var group in symbolGroups)
        {
            string symbol = group.Key;
            Security? security = Data[symbol];
            foreach (var cell in group)
            {
                IObserver<object> observer = cell.Key;
                if (security is not null)
                    observer.OnNext(ExcelFormatter.Get(security, cell.property));
                else
                {
                    Observers.Remove(observer);
                    string msg = $"YahooXL: symbol not found";
                    if (!symbol.StartsWith(msg))
                        msg += ": \"{symbol}\"";
                    observer.OnNext(msg + ".");
                    observer.OnCompleted();
                }
            }
        }
    }
}
