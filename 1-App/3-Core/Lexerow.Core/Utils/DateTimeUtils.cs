namespace Lexerow.Core.Utils;

public class DateTimeUtils
{
    public static TimeOnly ToTimeOnly(double val)
    {
        DateTime dtVal = DateTime.FromOADate(val);
        return new TimeOnly(dtVal.Hour, dtVal.Minute, dtVal.Second, dtVal.Millisecond);
    }

    public static DateTime ToDateTime(double val)
    {
        return DateTime.FromOADate(val);
    }
}