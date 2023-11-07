namespace YahooXL;

public static class ExcelHelpers
{
    private static readonly DateTimeZone SystemTimeZone = DateTimeZoneProviders.Tzdb.GetSystemDefault();

    // Default timeZone is system, else use "UTC", "America/New_York"...
    [ExcelFunction(Description = "DateTime Helper")] 
    public static object UnixSecondsToDateTime(long unixSeconds, string timeZoneName = "")
    {
        if (unixSeconds <= 0)
            return "Invalid unixSeconds.";
        Instant instant = Instant.FromUnixTimeSeconds(unixSeconds);
        DateTimeZone? tz = string.IsNullOrWhiteSpace(timeZoneName) ? SystemTimeZone : DateTimeZoneProviders.Tzdb.GetZoneOrNull(timeZoneName);
        if (tz is null)
            return $"Invalid timeZone: {timeZoneName}.";
        return instant.InZone(tz).ToDateTimeUnspecified().ToOADate();
    }

    // Default timeZone is system, else use "UTC", "America/New_York"...
    [ExcelFunction(Description = "DateTime Helper")]
    public static object UnixMillisecondsToDateTime(long unixMilliseconds, string timeZoneName = "")
    {
        if (unixMilliseconds <= 0)
            return "Invalid unixMilliseconds.";
        Instant instant = Instant.FromUnixTimeMilliseconds(unixMilliseconds);
        DateTimeZone? tz = string.IsNullOrWhiteSpace(timeZoneName) ? SystemTimeZone : DateTimeZoneProviders.Tzdb.GetZoneOrNull(timeZoneName);
        if (tz is null)
            return $"Invalid timeZone: {timeZoneName}.";
        return instant.InZone(tz).ToDateTimeUnspecified().ToOADate();
    }
}
