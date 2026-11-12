// The members are bound by name from scripts (so they must stay instance members) and the names are those of the WMI class
#pragma warning disable CA1822
using System;
using System.Linq;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The Win32_TimeZone instance of the local time zone (TimeZoneInfo.Local), with the properties and types that Windows WMI has (issue #101):
    /// Bias is the offset from UTC in minutes (60 for UTC+01:00 - verified with cscript), StandardBias / DaylightBias the bias of the standard and
    /// daylight saving time relative to it (0 and -60 for a one hour shift), the Standard* / Daylight* transition values of the current rule
    /// (Day is the week of the month, 5 the last one; DayOfWeek a Byte, 0 is Sunday) - all 0 for a time zone without daylight saving time
    /// </summary>
    [SourceClassName("SWbemObjectEx")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWin32TimeZone
    {
        public const string ClassName = "Win32_TimeZone";

        private ScriptonWin32TimeZone()
        {
        }

        public static ScriptonWin32TimeZone FromLocal() => From(TimeZoneInfo.Local, DateTime.Now);

        internal static ScriptonWin32TimeZone From(TimeZoneInfo timeZone, DateTime now)
        {
            var instance = new ScriptonWin32TimeZone
            {
                Bias = (int)timeZone.BaseUtcOffset.TotalMinutes,
                Caption = timeZone.DisplayName,
                Description = timeZone.DisplayName,
                StandardName = timeZone.StandardName,
                DaylightName = timeZone.DaylightName,
            };
            TimeZoneInfo.AdjustmentRule? rule = timeZone.SupportsDaylightSavingTime
                ? timeZone.GetAdjustmentRules().FirstOrDefault(r => (r.DateStart <= now.Date) && (now.Date <= r.DateEnd)) ?? timeZone.GetAdjustmentRules().LastOrDefault()
                : null;
            if ((rule != null) && (rule.DaylightDelta != TimeSpan.Zero))
            {
                instance.DaylightBias = -(int)rule.DaylightDelta.TotalMinutes;
                SetTransition(rule.DaylightTransitionStart, out instance._daylight);
                SetTransition(rule.DaylightTransitionEnd, out instance._standard);
            }
            return instance;
        }

        private struct Transition
        {
            public int Month;
            public int Day;
            public byte DayOfWeek;
            public int Hour;
            public int Minute;
            public int Second;
            public int Millisecond;
        }

        private Transition _daylight;
        private Transition _standard;

        private static void SetTransition(TimeZoneInfo.TransitionTime time, out Transition transition)
        {
            transition = new Transition
            {
                Month = time.Month,
                Day = time.IsFixedDateRule ? time.Day : time.Week,
                DayOfWeek = time.IsFixedDateRule ? (byte)0 : (byte)time.DayOfWeek,
                Hour = time.TimeOfDay.Hour,
                Minute = time.TimeOfDay.Minute,
                Second = time.TimeOfDay.Second,
                Millisecond = time.TimeOfDay.Millisecond,
            };
        }

        public int Bias { get; private set; }
        public string Caption { get; private set; } = "";
        public string Description { get; private set; } = "";
        public object SettingID => DBNull.Value;

        public string StandardName { get; private set; } = "";
        public int StandardBias => 0;
        public int StandardYear => 0;
        public int StandardMonth => _standard.Month;
        public int StandardDay => _standard.Day;
        public byte StandardDayOfWeek => _standard.DayOfWeek;
        public int StandardHour => _standard.Hour;
        public int StandardMinute => _standard.Minute;
        public int StandardSecond => _standard.Second;
        public int StandardMillisecond => _standard.Millisecond;

        public string DaylightName { get; private set; } = "";
        public int DaylightBias { get; private set; }
        public int DaylightYear => 0;
        public int DaylightMonth => _daylight.Month;
        public int DaylightDay => _daylight.Day;
        public byte DaylightDayOfWeek => _daylight.DayOfWeek;
        public int DaylightHour => _daylight.Hour;
        public int DaylightMinute => _daylight.Minute;
        public int DaylightSecond => _daylight.Second;
        public int DaylightMillisecond => _daylight.Millisecond;
    }
}
