using System;

namespace SheetStream.Internal
{
    /// <summary>
    /// Converts between <see cref="DateTime"/> and the serial numbers Excel stores in the 1900 date system.
    /// </summary>
    /// <remarks>
    /// Excel inherited from Lotus 1-2-3 the belief that 1900 was a leap year, so it has a serial (60) for
    /// 29 February 1900, a day that never existed. <see cref="DateTime.ToOADate"/> does not share the bug:
    /// from 1 March 1900 on the two agree, but before it the OLE date is one higher than Excel's serial
    /// (1 January 1900 is OLE date 2 and Excel serial 1). Writing the OLE date unchanged made January and
    /// February 1900 show one day late in Excel. Dates before 1 January 1900 cannot be represented in Excel at
    /// all and are passed through unchanged, as are time-only values (a serial below 1).
    /// </remarks>
    internal static class ExcelDate
    {
        /// <summary>The serial Excel uses for <paramref name="value"/>.</summary>
        public static double ToSerial(DateTime value)
        {
            var oa = value.ToOADate();
            return oa >= 2 && oa < 61 ? oa - 1 : oa;
        }

        /// <summary>
        /// The date an Excel serial stands for. Serial 60, Excel's fictitious 29 February 1900, reads as
        /// 28 February because there is no real day to map it to.
        /// </summary>
        public static DateTime FromSerial(double serial)
            => DateTime.FromOADate(serial >= 1 && serial < 60 ? serial + 1 : serial);
    }
}
