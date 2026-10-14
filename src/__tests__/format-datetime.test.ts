import { describe, it, expect } from 'vitest';
import { formatVbDateTime, VbDateTimeFormat } from '../runtime/vb-date.ts';
import { LCID_TO_LOCALE } from '../builtins/locale-data.generated.ts';

// FormatDateTime against the real vbscript.dll.
//
// Every expected value below was measured on this machine with `cscript //E:vbscript`, one fresh process
// per locale (a single process that calls SetLocale in a loop is not reliable: the values it reports for
// some locales are not the ones a fresh process reports). Each case is the instant 2026-01-05 09:05:00
// unless the row says otherwise, and the formats are vbGeneralDate..vbShortTime, which are 0..4.
//
// The engine's own values come from the Windows NLS patterns shipped with it, not from CLDR: for the ten
// locales below Intl disagrees with vbscript.dll on at least one format, which is why the table exists.
const ENGINE_LOCALE = (lcid: number): string => LCID_TO_LOCALE.get(lcid) ?? 'en-us';

/** The measured value of one FormatDateTime call, keyed by LCID and named format. */
const MEASURED: Record<number, Record<number, string>> = {
  // de-DE: day and month padded, 24-hour clock, long date with a comma.
  1031: {
    0: '05.01.2026 09:05:00',
    1: 'Montag, 5. Januar 2026',
    2: '05.01.2026',
    3: '09:05:00',
    4: '09:05',
  },
  // en-US: the one common locale with an unpadded short date and a 12-hour clock.
  1033: {
    0: '1/5/2026 9:05:00 AM',
    1: 'Monday, January 5, 2026',
    2: '1/5/2026',
    3: '9:05:00 AM',
    4: '09:05',
  },
  // zh-CN: ideographic long date, unpadded numeric date, 24-hour clock with no leading zero.
  2052: {
    0: '2026/1/5 9:05:00',
    1: '2026年1月5日',
    2: '2026/1/5',
    3: '9:05:00',
    4: '09:05',
  },
  // ko-KR: the clock carries its AM/PM designator first, the long date its weekday last.
  1042: {
    0: '2026-01-05 오전 9:05:00',
    1: '2026년 1월 5일 월요일',
    2: '2026-01-05',
    3: '오전 9:05:00',
    4: '09:05',
  },
  // ar-SA, measured in a fresh process: the two date formats are Hijri and NOT the Gregorian date CStr
  // prints, while GeneralDate, the long time and the short time are Gregorian. The Hijri short date also
  // shows that a Windows "yyyy" row writes two year digits here ("16/07/47", not "16/07/1447").
  //
  // The long date is the whole Hijri date with the calendar's own month name. Windows separates its three
  // fields with "/" where the table's picture string writes a space; the engine keeps the picture string's
  // separator. That is the documented deviation for the RTL locales (see the note above MEASURED).
  1025: {
    0: '05/01/2026 09:05:00 ص',
    1: '‏16 ‏رجب, ‏1447',
    2: '‏16/‏07/‏47',
    3: '09:05:00 ص',
    4: '09:05',
  },
  // sv-SE: a long date with a word in front of the day.
  1053: {
    0: '2026-01-05 09:05:00',
    1: 'den 5 januari 2026',
    2: '2026-01-05',
    3: '09:05:00',
    4: '09:05',
  },
  // en-GB: day before month, no leading comma in the long date.
  2057: {
    0: '05/01/2026 09:05:00',
    1: '05 January 2026',
    2: '05/01/2026',
    3: '09:05:00',
    4: '09:05',
  },
  // hi-IN: a dash separates the short date where most locales use a slash or a dot.
  1081: {
    0: '05-01-2026 09:05:00',
    1: '05 जनवरी 2026',
    2: '05-01-2026',
    3: '09:05:00',
    4: '09:05',
  },
  // fr-FR: an unpadded day, and a long date with no weekday at all.
  1036: {
    0: '05/01/2026 09:05:00',
    1: 'lundi 5 janvier 2026',
    2: '05/01/2026',
    3: '09:05:00',
    4: '09:05',
  },
  // es-ES with the traditional sort order: "de" between the fields, and a single-digit long-time hour.
  1034: {
    0: '05/01/2026 9:05:00',
    1: 'lunes, 05 de enero de 2026',
    2: '05/01/2026',
    3: '9:05:00',
    4: '09:05',
  },
};

const instant = (hour: number, minute: number): Date => new Date(2026, 0, 5, hour, minute, 0);

describe('FormatDateTime', () => {
  for (const [lcidText, formats] of Object.entries(MEASURED)) {
    const lcid = Number(lcidText);
    it(`matches vbscript.dll for every named format under ${lcid}`, () => {
      for (const [styleText, expected] of Object.entries(formats)) {
        const style = Number(styleText);
        expect(
          formatVbDateTime(instant(9, 5), lcid, style, ENGINE_LOCALE(lcid)),
          `LCID ${lcid} (${ENGINE_LOCALE(lcid)}), format ${style}`
        ).toBe(expected);
      }
    });
  }

  it('names the formats as the OLE Automation constants do', () => {
    expect(VbDateTimeFormat).toEqual({ GeneralDate: 0, LongDate: 1, ShortDate: 2, LongTime: 3, ShortTime: 4 });
  });

  // Measured on 1031 and 1033: midnight of an ordinary day gives the date alone for GeneralDate, while
  // the three date formats are unchanged and the two clocks still read as a clock.
  it('drops the time from GeneralDate at midnight, as vbscript.dll does', () => {
    expect(formatVbDateTime(instant(0, 0), 1031, VbDateTimeFormat.GeneralDate, 'de-de')).toBe('05.01.2026');
    expect(formatVbDateTime(instant(0, 0), 1033, VbDateTimeFormat.GeneralDate, 'en-us')).toBe('1/5/2026');
    expect(formatVbDateTime(instant(0, 0), 1033, VbDateTimeFormat.LongTime, 'en-us')).toBe('12:00:00 AM');
    expect(formatVbDateTime(instant(0, 0), 1033, VbDateTimeFormat.ShortTime, 'en-us')).toBe('00:00');
  });

  // Measured: style 4 is the locale's own LOCALE_SSHORTTIME with the designator gone and a padded hour,
  // not a hard-coded "HH:mm". Finnish and Indonesian keep their own separator, Korean drops its "오전".
  it('keeps the locale separator in the short time and pads the hour', () => {
    expect(formatVbDateTime(instant(9, 5), 1035, VbDateTimeFormat.ShortTime, 'fi-fi')).toBe('09.05');
    expect(formatVbDateTime(instant(9, 5), 1057, VbDateTimeFormat.ShortTime, 'id-id')).toBe('09.05');
    expect(formatVbDateTime(instant(9, 5), 1042, VbDateTimeFormat.ShortTime, 'ko-kr')).toBe('09:05');
    expect(formatVbDateTime(instant(9, 5), 1026, VbDateTimeFormat.ShortTime, 'bg-bg')).toBe('09:05');
  });

  // Measured: Thai is the one locale of the 131 the table covers where Windows has no date format at all,
  // and it raises error 5 (invalid procedure call) for LongDate and ShortDate on every instant. The engine
  // renders those two formats instead, from the Buddhist calendar Intl supplies - a deliberate deviation,
  // recorded here so it is visible. Its two clocks and its GeneralDate match Windows exactly.
  it('renders the Thai dates Windows refuses to write, and matches it on the clocks', () => {
    expect(formatVbDateTime(instant(9, 5), 1054, VbDateTimeFormat.LongDate, 'th-th')).toBe(
      'วันจันทร์ที่ 5 มกราคม 2569'
    );
    expect(formatVbDateTime(instant(9, 5), 1054, VbDateTimeFormat.ShortDate, 'th-th')).toBe('5/1/2569');
    expect(formatVbDateTime(instant(9, 5), 1054, VbDateTimeFormat.GeneralDate, 'th-th')).toBe('5/1/2026 9:05:00');
    expect(formatVbDateTime(instant(9, 5), 1054, VbDateTimeFormat.LongTime, 'th-th')).toBe('9:05:00');
    expect(formatVbDateTime(instant(9, 5), 1054, VbDateTimeFormat.ShortTime, 'th-th')).toBe('09:05');
  });

  // Measured: on day 0 every locale prints the time alone for GeneralDate, exactly as CStr does, except a
  // Hijri one, which keeps its date (ar-SA prints "30/12/1899 09:05:00 ص", the Gregorian short date).
  it('prints the time alone on day 0 for GeneralDate, except in the Hijri locales', () => {
    const dayZero = new Date(1899, 11, 30, 9, 5, 0);
    expect(formatVbDateTime(dayZero, 1026, VbDateTimeFormat.GeneralDate, 'bg-bg')).toBe('9:05:00');
    expect(formatVbDateTime(dayZero, 1025, VbDateTimeFormat.GeneralDate, 'ar-sa')).toBe('30/12/1899 09:05:00 ص');
  });

  it('keeps the whole date at midnight for the date formats, unlike GeneralDate', () => {
    expect(formatVbDateTime(instant(0, 0), 1031, VbDateTimeFormat.LongDate, 'de-de')).toBe('Montag, 5. Januar 2026');
    expect(formatVbDateTime(instant(0, 0), 1031, VbDateTimeFormat.ShortDate, 'de-de')).toBe('05.01.2026');
  });

  it('writes the Thai long date with the short weekday, which its picture string asks for', () => {
    expect(formatVbDateTime(instant(9, 5), 1054, VbDateTimeFormat.LongDate, 'th-th')).toBe(
      'วันจันทร์ที่ 5 มกราคม 2569'
    );
  });
});
