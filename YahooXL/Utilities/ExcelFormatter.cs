using System.Reflection;
using NodaTime.Text;

namespace YahooXL;

public static class ExcelFormatter
{
    private static readonly LocalDatePattern Ldp = LocalDatePattern.CreateWithInvariantCulture("yyyy-MM-dd");
    private static readonly LocalTimePattern Ltp = LocalTimePattern.CreateWithInvariantCulture("HH:mm:ss");
    private static readonly LocalDateTimePattern Ldtp = LocalDateTimePattern.CreateWithInvariantCulture("yyyy-MM-dd HH:mm:ss");

    internal static object Get(Security security, string property)
    {
        PropertyInfo pi = typeof(Security).GetProperty(property) ?? throw new InvalidOperationException();
        object? v = pi.GetValue(security);
        if (v is null)
            return "";
        if (v is Symbol s)
            return s.Name;
        if (v is ZonedDateTime zdt)
        {
            if (zdt == default)
                return "";
            return zdt.ToString();
        }
        if (v is LocalDateTime ldt)
        {
            if (ldt == default)
                return "";
            if (ldt.TickOfDay == 0)
                return Ldp.Format(ldt.Date);
            return Ldtp.Format(ldt);
        }
        if (v is LocalDate ld)
        {
            if (ld == default)
                return "";
            return Ldp.Format(ld);
        }
        if (v is LocalTime lt)
        {
            if (lt == default)
                return "";
            return Ltp.Format(lt);
        }
        return v;
    }
}
