using System.Collections.Concurrent;
using System.Reactive.Concurrency;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using Microsoft.Extensions.Logging;

namespace YahooXL;

public static class YahooQuotesAddin
{
    internal static ILogger Logger = AddIn.LogFactory.CreateLogger("YahooXL");
    internal static IDisposable RefreshDisposable = Disposable.Empty;
    private static readonly Dictionary<IObserver<object>, (string symbol, string property)> Observers = [];
    private static readonly BlockingCollection<(IObserver<object> observer, string symbol, string property)> ObserversToAdd = [];
    private static readonly ConcurrentBag<IObserver<object>> concurrentBag = [];
    private static readonly ConcurrentBag<IObserver<object>> ObserversToRemove = concurrentBag;
    private static Dictionary<string, Security?> Data = new(StringComparer.OrdinalIgnoreCase);
    private static int started = 0;

    [ExcelFunction(Description = "YahooXL: Delayed Quotes")]
    public static IObservable<object> YahooQuote([ExcelArgument("Symbol")] string symbol, [ExcelArgument("Property")] string property)
    {
        //ExcelReference? caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
        if (string.IsNullOrEmpty(symbol) && string.IsNullOrEmpty(property))
            return Observable.Return(Assembly.GetExecutingAssembly().GetName().ToString());
        Logger.LogDebug("YahooQuote({Symbol}, {Property})", symbol, property);
        if (string.IsNullOrEmpty(property))
            property = "RegularMarketPrice";
        else if (!Securities.PropertyExists(property) && int.TryParse(property, out int i))
        {
            string? name = Securities.GetPropertyNameByIndex(i);
            if (name != null)
                return Observable.Return(name);
            if (string.IsNullOrEmpty(symbol))
                return Observable.Return("YahooXL: Invalid property index.");
        }
        if (string.IsNullOrWhiteSpace(symbol))
            return Observable.Return(Securities.PropertyExists(property) ? property : "YahooXL: Unknown property.");

        object? value = "~";

        // symbol has already been requested
        if (Data.TryGetValue(symbol, out Security? security))
        {
            if (security is null)
            {
                string msg = "YahooXL: symbol not found";
                if (!symbol.StartsWith(msg))
                    msg += $": \"{symbol}\"";
                return Observable.Return(msg + ".");
            }
            value = Securities.GetValue(security, property);
        }

        return Observable.Create<object>(observer => // executed once
        {
            try
            {
                observer.OnNext(value ?? "YahooXL: null value.");
                ObserversToAdd.Add((observer, symbol, property));
                if (Interlocked.CompareExchange(ref started, 1, 0) == 0)
                    RefreshDisposable = NewThreadScheduler.Default.ScheduleAsync(RefreshLoop);
            }
            catch (Exception e)
            {
                Logger.LogError(e, "Could not add observer!");
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
            Logger.LogDebug("Loop starting.");
            while (true)
            {
                int wait = AddIn.Options.RtdIntervalMaxSeconds;
                while (true)
                {
                    if (!ObserversToAdd.TryTake(out (IObserver<object> observer, string symbol, string property) Item, wait * 1000, ct))
                        break;
                    wait = AddIn.Options.RtdIntervalMinSeconds;
                    Observers.Add(Item.observer, (Item.symbol, Item.property));
                    Logger.LogTrace("Added observer: {Symbol}, {Property}.", Item.symbol, Item.property);
                }

                while (ObserversToRemove.TryTake(out IObserver<object>? observer))
                {
                    Observers.Remove(observer);
                    Logger.LogTrace("Removed observer.");
                }

                if (Observers.Count != 0)
                    await Update(ct).ConfigureAwait(false);
            }
        }
        catch (Exception e)
        {
            Logger.LogError(e, "Unexpected exception!");
        }
        finally
        {
            Logger.LogDebug("Loop ending.");
            started = 0;
        }
    }

    private static async Task Update(CancellationToken ct)
    {
        Logger.LogDebug("Updating.");

        var symbolGroups = Observers
            .Select(kvp => (kvp.Value.symbol, kvp.Value.property, kvp.Key))
            .GroupBy(x => x.symbol, StringComparer.OrdinalIgnoreCase)
            .ToList();

        IEnumerable<string> symbols = symbolGroups.Select(g => g.Key);

        try
        {
            // Reference assignment is guaranteed to be atomic.
            Data = await YahooFinance.GetSecuritiesAsync(symbols, ct).ConfigureAwait(false);
        } 
        catch (Exception e) 
        {
            Logger.LogError(e, "Yahoo.Update()");
            return;
        }

        foreach (var group in symbolGroups)
        {
            string symbol = group.Key;
            Security? security = Data[symbol];
            foreach (var cell in group)
            {
                IObserver<object> observer = cell.Key;
                if (security is null)
                {
                    Observers.Remove(observer);
                    string msg = "YahooXL: symbol not found";
                    if (!symbol.StartsWith(msg))
                        msg += $": \"{symbol}\"";
                    observer.OnNext(msg + ".");
                    observer.OnCompleted();
                }
                else
                {
                    observer.OnNext(Securities.GetValue(security, cell.property));
                    //observer.OnNext(DateTime.Now.ToOADate());
                }
            }
        }
    }
}
