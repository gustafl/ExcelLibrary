namespace ExcelLibrary;

/// <summary>
/// Internal utility methods for data conversion.
/// </summary>
static class Utilities
{
    /// <summary>
    /// Excel's epoch starts at day 1, but DateOnly counts from day 0, and Excel incorrectly
    /// treats 1900 as a leap year (the "Lotus 1-2-3 bug"), so we subtract 2 days.
    /// </summary>
    private const int ExcelEpochOffset = 2;

    /// <summary>
    /// Number of seconds in a day (60 * 60 * 24).
    /// </summary>
    private const int SecondsPerDay = 86400;

    /// <summary>
    /// Converts an Excel serial date number to a formatted date string.
    /// </summary>
    /// <param name="excelDate">The Excel serial date number as a string.</param>
    /// <param name="baseYear">The base year (1900 or 1904) for date calculations.</param>
    /// <returns>A formatted date string.</returns>
    internal static string ConvertDate(string excelDate, int baseYear)
    {
        var baseDate = new DateOnly(baseYear, 1, 1);
        int daysToAdd = (int)(double.Parse(excelDate) - ExcelEpochOffset);
        var convertedDate = baseDate.AddDays(daysToAdd);
        return convertedDate.ToShortDateString();
    }

    /// <summary>
    /// Converts an Excel time fraction to a formatted time string.
    /// </summary>
    /// <param name="excelTime">The Excel time as a decimal fraction of a day.</param>
    /// <returns>A formatted time string (HH:mm:ss for times under 24 hours, TimeSpan format otherwise).</returns>
    internal static string ConvertTime(string excelTime)
    {
        double time = double.Parse(excelTime, CultureInfo.GetCultureInfo("en-us"));
        double seconds = time * SecondsPerDay;

        // Use TimeOnly for times within 24 hours, otherwise fall back to TimeSpan
        if (seconds is >= 0 and < SecondsPerDay)
        {
            var timeOnly = TimeOnly.FromTimeSpan(TimeSpan.FromSeconds(seconds));
            return timeOnly.ToString("HH:mm:ss");
        }

        return TimeSpan.FromSeconds(seconds).ToString();
    }
}
