using System.Reactive.Linq;
using ExcelDna.Registration;
namespace YahooXL;

public static class SampleFunctions
{

#if (DEBUG)
    [ExcelFunction]
    public static object Dna() => DateTime.Now.ToString("HH:mm:ss");

    [ExcelAsyncFunction]
    public static async Task<string> DnaAsync()
    {
        await Task.CompletedTask;
        return DateTime.Now.ToString("HH:mm:ss");
    }

    [ExcelFunction]
    public static IObservable<string> DnaObservableComplete()
    {
        return Observable.Return(DateTime.Now.ToString("HH:mm:ss"));
    }

    [ExcelFunction]
    public static IObservable<string> DnaObservableInComplete() // will update when the worksheet opens
    {
        return YahooQuotesAddin.ObservableReturn(DateTime.Now.ToString("HH:mm:ss"));
        //return Observable.Never<string>().StartWith(DateTime.Now.ToString("HH:mm:ss");
    }

    [ExcelFunction]
    public static object DnaArray()
    {
        return new object[,] { { 1, DateTime.Now.ToString("HH:mm:ss") }, { 2, DateTime.Now.ToString("HH:mm:ss") } };
    }

    [ExcelFunction]
    public static IObservable<object> DnaObservableArray()
    {
        return Observable.Timer(dueTime: TimeSpan.Zero, period: TimeSpan.FromSeconds(1))
        .Select(_ => new object[,] { { 1, DateTime.Now.ToString("HH:mm:ss") }, { 2, DateTime.Now.ToString("HH:mm:ss") } });
    }

#endif

    // Govert: I think the best way to make an RTD server is to make a Rx IObservable,
    // and then exposing with or without the Registration helper library.
    // The ExcelDna.Registration extension library, which will automatically generate the
    // wrapper function from your function returning IObservable<T>. However in this case
    // you need to add the explicit registration processing code into your AutoOpen.
    [ExcelFunction(IsThreadSafe = true)]
    public static IObservable<string> DnaObservableClock()
    {
        return Observable.Timer(dueTime: TimeSpan.Zero, period: TimeSpan.FromSeconds(1))
            .Select(_ => DateTime.Now.ToString("HH:mm:ss"));
    }


}
