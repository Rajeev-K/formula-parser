using System;
using System.Text;
using System.Globalization;
using Microsoft.VisualBasic;
using System.Reflection;

// For security reasons, formulas are only allowed to call functions defined in this file.

namespace FormulaParser
{
    [Obfuscation(Exclude = true, ApplyToMembers = true)]
    public enum DateInterval
    {
        Year = 0, Quarter = 1, Month = 2, DayOfYear = 3, Day = 4,
        WeekOfYear = 5, Weekday = 6, Hour = 7, Minute = 8, Second = 9
    }

    [Obfuscation(Exclude = true, ApplyToMembers = true)]
    public enum FirstDayOfWeek
    {
        System = 0, Sunday = 1, Monday = 2, Tuesday = 3,
        Wednesday = 4, Thursday = 5, Friday = 6, Saturday = 7
    }

    [Obfuscation(Exclude = true, ApplyToMembers = true)]
    public class Functions
    {
        // --- Control flow ---

        public static object Iif(bool condition, object a, object b)
        {
            return condition ? a : b;
        }

        public static bool IsNothing(object a)
        {
            return a == null;
        }

        // --- Text manipulation ---

        public static string Format(object expression, string style)
        {
            return Strings.Format(expression, style);
        }

        public static int Len(string str)
        {
            return Strings.Len(str);
        }

        public static int StrComp(string String1, string String2)
        {
            return Strings.StrComp(String1, String2, CompareMethod.Text);
        }

        public static string Replace(string Expression, string Find, string Replacement)
        {
            return Strings.Replace(Expression, Find, Replacement, 1, -1, CompareMethod.Text);
        }

        // --- index of substring ---

        public static int InStr(string StringCheck, string StringSearch)
        {
            return Strings.InStr(StringCheck, StringSearch, CompareMethod.Binary);
        }

        public static int InStrRev(string StringCheck, string StringSearch)
        {
            return Strings.InStrRev(StringCheck, StringSearch, -1, CompareMethod.Binary);
        }

        // --- Substring ---

        public static string Left(string str, int Length)
        {
            return Strings.Left(str, Length);
        }

        public static string Right(string str, int Length)
        {
            return Strings.Right(str, Length);
        }

        public static string Mid(string str, int Start, int Length)
        {
            return Strings.Mid(str, Start, Length);
        }

        public static char GetChar(string str, int index)
        {
            return Strings.GetChar(str, index);
        }

        // --- Trimming ---

        public static string LTrim(string str)
        {
            return Strings.LTrim(str);
        }

        public static string RTrim(string str)
        {
            return Strings.RTrim(str);
        }

        public static string Trim(string str)
        {
            return Strings.Trim(str);
        }

        // --- Case conversion ---

        public static string LCase(string value)
        {
            return Strings.LCase(value);
        }

        public static string UCase(string value)
        {
            return Strings.UCase(value);
        }

        public static string TitleCase(string s)
        {
            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s);
        }

        // --- Date manipulation ---

        public static DateTime DateAdd(DateInterval Interval, double Number, DateTime DateValue)
        {
            return DateAndTime.DateAdd((Microsoft.VisualBasic.DateInterval)Interval, Number, DateValue);
        }

        public static long DateDiff(DateInterval Interval, DateTime Date1, DateTime Date2) 
        {
            return DateAndTime.DateDiff((Microsoft.VisualBasic.DateInterval)Interval, Date1, Date2,
                            Microsoft.VisualBasic.FirstDayOfWeek.Sunday, Microsoft.VisualBasic.FirstWeekOfYear.Jan1);
        }

        public static int DatePart(DateInterval Interval, DateTime DateValue)
        {
            return DateAndTime.DatePart((Microsoft.VisualBasic.DateInterval)Interval, DateValue,
                           Microsoft.VisualBasic.FirstDayOfWeek.Sunday, Microsoft.VisualBasic.FirstWeekOfYear.Jan1);
        }

        public static DateTime DateSerial(int Year, int Month, int Day)
        {
            return DateAndTime.DateSerial(Year, Month, Day);
        }

        public static DateTime DateValue(string StringDate)
        {
            return DateAndTime.DateValue(StringDate);
        }

        public static int Day(DateTime DateValue)
        {
            return DateAndTime.Day(DateValue);
        }

        public static int Hour(DateTime TimeValue)
        {
            return DateAndTime.Hour(TimeValue);
        }

        public static int Minute(DateTime TimeValue)
        {
            return DateAndTime.Minute(TimeValue);
        }

        public static int Month(DateTime DateValue)
        {
            return DateAndTime.Month(DateValue);
        }

        public static string MonthName(int Month, bool Abbreviate)
        {
            return DateAndTime.MonthName(Month, Abbreviate);
        }

        public static int Second(DateTime TimeValue)
        {
            return DateAndTime.Second(TimeValue);
        }

        public static DateTime TimeSerial(int Hour, int Minute, int Second)
        {
            return DateAndTime.TimeSerial(Hour, Minute, Second);
        }

        public static DateTime TimeValue(string StringTime)
        {
            return DateAndTime.TimeValue(StringTime);
        }

        public static int Weekday(DateTime DateValue, FirstDayOfWeek DayOfWeek)
        {
            return DateAndTime.Weekday(DateValue, (Microsoft.VisualBasic.FirstDayOfWeek)DayOfWeek);
        }

        public static string WeekdayName(int Weekday, bool Abbreviate, FirstDayOfWeek FirstDayOfWeekValue)
        {
            return DateAndTime.WeekdayName(Weekday, Abbreviate, (Microsoft.VisualBasic.FirstDayOfWeek)FirstDayOfWeekValue);
        }

        public static int Year(DateTime DateValue)
        {
            return DateAndTime.Year(DateValue);
        }

        public static DateTime Now
        {
            get { return DateAndTime.Now; }
        }

        public static DateTime Today
        {
            get { return DateAndTime.Today; }
        }

        public static DateTime TimeOfDay
        {
            get { return DateAndTime.TimeOfDay; }
        }

        // --- Type conversion ---

        public static double Fix(double Number)
        {
            return Conversion.Fix(Number);
        }

        public static double Int(double Number)
        {
            return Conversion.Int(Number);
        }

        public static string Str(object Number)
        {
            return Conversion.Str(Number);
        }

        public static double Val(string InputStr)
        {
            return Conversion.Val(InputStr);
        }

        public static bool CBool(object obj)
        {
            return Convert.ToBoolean(obj);
        }

        public static DateTime CDate(object obj)
        {
            return Convert.ToDateTime(obj);
        }

        public static double CDbl(object obj)
        {
            return Convert.ToDouble(obj);
        }

        public static int CInt(object obj)
        {
            return Convert.ToInt32(obj);
        }

        public static string CStr(object obj)
        {
            return Convert.ToString(obj);
        }

        // --- Math ---

        public static double Abs(double val)
        {
            return Math.Abs(val);
        }

        public static int Abs(int val)
        {
            return Math.Abs(val);
        }

        public static double Ceiling(double d)
        {
            return Math.Ceiling(d);
        }

        public static double Floor(double d)
        {
            return Math.Floor(d);
        }

        public static double Max(double a, double b)
        {
            return Math.Max(a, b);
        }

        public static int Max(int a, int b)
        {
            return Math.Max(a, b);
        }

        public static double Min(double a, double b)
        {
            return Math.Min(a, b);
        }

        public static int Min(int a, int b)
        {
            return Math.Min(a, b);
        }

        public static double Pow(double x, double y)
        {
            return Math.Pow(x, y);
        }

        public static double Round(double a)
        {
            return Math.Round(a);
        }

        public static double Round(double value, int digits)
        {
            return Math.Round(value, digits);
        }

        public static int Sign(double value)
        {
            return Math.Sign(value);
        }

        public static int Sign(int value)
        {
            return Math.Sign(value);
        }

        public static double Sqrt(double d)
        {
            return Math.Sqrt(d);
        }

        public static double Truncate(double d)
        {
            return Math.Truncate(d);
        }

        public static double Log(double d)
        {
            return Math.Log(d);
        }

        public static double Log10(double d)
        {
            return Math.Log10(d);
        }

        public static double Sin(double a)
        {
            return Math.Sin(a);
        }

        public static double Cos(double a)
        {
            return Math.Cos(a);
        }

        public static double Tan(double a)
        {
            return Math.Tan(a);
        }

        public static double Asin(double a)
        {
            return Math.Asin(a);
        }

        public static double Acos(double a)
        {
            return Math.Acos(a);
        }

        public static double Atan(double a)
        {
            return Math.Atan(a);
        }

        public static double Atan2(double y, double x)
        {
            return Math.Atan2(y, x);
        }
    }
}
