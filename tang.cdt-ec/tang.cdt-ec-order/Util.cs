using System;

namespace tang.cdt_ec_order
{
    public class Util
    {
        public static DateTime ConvertToDateTime(string timeStamp)
        {
            DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1));

            var toConvertValue = long.Parse(timeStamp) / 1000;

            return startTime.AddSeconds(toConvertValue);
        }
    }
}
