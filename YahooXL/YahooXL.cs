using System.Collections.Concurrent;
using System.Reactive.Concurrency;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using Microsoft.Extensions.Logging;

namespace YahooXL;

public static class YahooQuotesAddin
{
    internal static readonly ILogger Logger = LoggerFactory
        .Create(builder => builder
            .AddDebug()
            .AddEventSourceLogger()
#if DEBUG
            .SetMinimumLevel(LogLevel.Debug))
#else
            .SetMinimumLevel(LogLevel.Warning))
#endif
        .CreateLogger("YahooXL");
    private static readonly IScheduler Scheduler = NewThreadScheduler.Default;
    private static readonly Dictionary<IObserver<object>, (string symbol, string property)> Observers = new();
    private static readonly BlockingCollection<(IObserver<object> observer, string symbol, string property)> ObserversToAdd = new();
    private static readonly ConcurrentBag<IObserver<object>> ObserversToRemove = new();
    private static Dictionary<string, Security?> Data = new(StringComparer.OrdinalIgnoreCase);
    private static int started = 0;

    [ExcelFunction(Description = "YahooXL: Delayed Quotes")]
    public static IObservable<object> YahooQuote([ExcelArgument("Symbol")] string symbol, [ExcelArgument("Property")] string property)
    {
        Logger.LogDebug("YahooQuote({Symbol}, {Property})", symbol, property);
        if (string.IsNullOrEmpty(symbol) && string.IsNullOrEmpty(property))
            return Observable.Return(Assembly.GetExecutingAssembly().GetName().ToString());
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
                    Scheduler.ScheduleAsync(RefreshLoop);
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
                int wait = 30000; // 30s
                while (true)
                {
                    if (!ObserversToAdd.TryTake(out (IObserver<object> observer, string symbol, string property) Item, wait, ct))
                        break;
                    wait = 1000; // 1s
                    Observers.Add(Item.observer, (Item.symbol, Item.property));
                    Logger.LogTrace("Added observer: {Symbol}, {Property}.", Item.symbol, Item.property);
                }

                while (ObserversToRemove.TryTake(out IObserver<object>? observer))
                {
                    Observers.Remove(observer);
                    Logger.LogTrace("Removed observer.");
                }
                Debug.Flush();

                if (Observers.Any())
                    await Update(ct).ConfigureAwait(false);
            }
        }
        catch (Exception e)
        {
            Logger.LogError(e, "Swallowed exception!");
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

        // Reference assignment is guaranteed to be atomic.
        Data = await Yahoo.GetSecuritiesAsync(symbols, ct).ConfigureAwait(false);

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
                    continue;
                }
                observer.OnNext(Securities.GetValue(security, cell.property));
            }
        }
    }
}
