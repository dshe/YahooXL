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
    private static readonly Dictionary<IObserver<object>, (Symbol symbol, string property)> Observers = [];
    private static readonly BlockingCollection<(IObserver<object> observer, Symbol symbol, string property)> ObserversToAdd = [];
    private static readonly ConcurrentBag<IObserver<object>> ObserversToRemove = [];
    private static Dictionary<Symbol, Snapshot?> Data = [];
    private static int started = 0;

    [ExcelFunction(Description = "YahooXL: Delayed Quotes", IsThreadSafe = true)]
    public static IObservable<object> YahooQuote(
        [ExcelArgument("Symbol")] string symbolArg, 
        [ExcelArgument("Property")] string propertyArg)
    {
        if (string.IsNullOrEmpty(symbolArg) && string.IsNullOrEmpty(propertyArg))
            return ObservableReturn(Assembly.GetExecutingAssembly().GetName().FullName);
        if (symbolArg == "#PathInfo")
            return ObservableReturn(ExcelDnaUtil.XllPathInfo.FullName);

        if (symbolArg.StartsWith('#') || symbolArg.StartsWith('~'))
            return Observable.Return(symbolArg);
        if (propertyArg.StartsWith('#'))
            return Observable.Return(propertyArg);

        Logger.LogDebug("YahooQuote('{Symbol}', '{Property}')", symbolArg, propertyArg);
        if (string.IsNullOrEmpty(propertyArg))
            propertyArg = "RegularMarketPrice";

        string property = YahooFinance.GetPropertyOrError(propertyArg);
        if (property.StartsWith('#') || string.IsNullOrWhiteSpace(symbolArg))
            return ObservableReturn(property);

        if (!Symbol.TryCreate(symbolArg, out Symbol symbol) || symbol.IsCurrency)
            return Observable.Return($"#Invalid symbol name: '{symbolArg}'.");

        object value = "~";

        if (Data.TryGetValue(symbol, out Snapshot? snapshot))
        {   // Symbol has already been requested
            if (snapshot is null)
                return Observable.Return($"#Symbol not found: '{symbol.Name}'.");
            value = snapshot.GetValue(property);
        }

        return Observable.Create<object>(observer => // executed once
        {
            try
            {
                observer.OnNext(value);
                ObserversToAdd.Add((observer, symbol, property));
                if (Interlocked.CompareExchange(ref started, 1, 0) == 0)
                    RefreshDisposable = NewThreadScheduler.Default.ScheduleAsync(RefreshLoop);
            }
            catch (Exception e)
            {
                Logger.LogError(e, "Could not add observer.");
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
                    if (!ObserversToAdd.TryTake(out (IObserver<object> observer, Symbol symbol, string property) Item, wait * 1000, ct))
                        break;
                    wait = AddIn.Options.RtdIntervalMinSeconds;
                    Observers.Add(Item.observer, (Item.symbol, Item.property));
                    Logger.LogTrace("Added observer: {Symbol}, {Property}.", Item.symbol.Name, Item.property);
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
            .GroupBy(x => x.symbol)
            .ToList();

        try
        {
            // Reference assignment is guaranteed to be atomic.
            IEnumerable<Symbol> symbols = symbolGroups.Select(g => g.Key);
            Data = await YahooFinance.GetSnapshotAsync(symbols, ct).ConfigureAwait(false);
        } 
        catch (Exception e) 
        {
            Logger.LogError(e, "Yahoo.Update()");
            return;
        }

        foreach (var group in symbolGroups)
        {
            Symbol symbol = group.Key;
            Snapshot? snapshot = Data[symbol];
            foreach (var cell in group)
            {
                IObserver<object> observer = cell.Key;
                if (snapshot is null)
                {
                    Observers.Remove(observer);
                    observer.OnNext($"#Symbol not found: '{symbol.Name}'.");
                    observer.OnCompleted();
                }
                else
                {
                    observer.OnNext(snapshot.GetValue(cell.property));
                }
            }
        }
    }

    // Observable.Return() sends one value, then completes.
    // ObservableReturn() sends one value but does not complete. The value will updates each time the worksheet opens.
    internal static IObservable<T> ObservableReturn<T>(T value)
    {
        return Observable.Create<T>(observer =>
        {
            observer.OnNext(value);
            //observer.OnCompleted();
            return Disposable.Empty;
        });
    }
}
