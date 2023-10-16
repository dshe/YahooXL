using ExcelDna.Integration;

namespace YahooXL;

public static class ExcelHelpers
{
    private static readonly DateTimeZone SystemTimeZone = DateTimeZoneProviders.Tzdb.GetSystemDefault();

    // Default is LocalTime, else use "UTC" or other tzdb name such as "America/New_York".
    [ExcelFunction(Description = "DateTime Helper")] 
    public static object UnixSecondsToDateTime(long unixSeconds, string timezoneName = "")
    {
        if (unixSeconds < 0)
            return "Invalid unixSeconds.";
        Instant instant = Instant.FromUnixTimeSeconds(unixSeconds); // utc
        DateTimeZone? tz = string.IsNullOrWhiteSpace(timezoneName) ? SystemTimeZone : DateTimeZoneProviders.Tzdb.GetZoneOrNull(timezoneName);
        if (tz is null)
            return $"Invalid timezone: {timezoneName}.";
        ZonedDateTime zdt = instant.InZone(tz);
        return zdt.ToDateTimeUnspecified().ToOADate();
    }
}
