namespace ExcelLibrary;

static class Utilities
{
    public static string ConvertDate(string excelDate, int baseYear)
    {
        const int MYSTERY_CONSTANT = 2;
        var baseDate = new DateOnly(baseYear, 1, 1);
        int daysToAdd = (int)(double.Parse(excelDate) - MYSTERY_CONSTANT);
        var convertedDate = baseDate.AddDays(daysToAdd);
        return convertedDate.ToShortDateString();
    }

    public static string ConvertTime(string excelTime)
    {
        double time = double.Parse(excelTime, CultureInfo.GetCultureInfo("en-us"));
        double seconds = time * 86400;

        // Use TimeOnly for times within 24 hours, otherwise fall back to TimeSpan
        if (seconds is >= 0 and < 86400)
        {
            var timeOnly = TimeOnly.FromTimeSpan(TimeSpan.FromSeconds(seconds));
            return timeOnly.ToString("HH:mm:ss");
        }

        return TimeSpan.FromSeconds(seconds).ToString();
    }
}
