// Windows NLS date/time patterns. This file is data read from Windows, not a build artifact, and editing
// it by hand makes it something other than what Windows answers.
//
// Read with GetLocaleInfoEx(locale, LOCALE_NOUSEROVERRIDE):
//   LOCALE_SSHORTDATE, LOCALE_SLONGDATE, LOCALE_STIMEFORMAT, LOCALE_SSHORTTIME,
//   LOCALE_S1159, LOCALE_S2359, LOCALE_ICALENDARTYPE
//
// Read on Windows build 26200 25H2, 2026-10-14
// Rows: lcid|shortDate|longDate|timeFormat|shortTime|am|pm|calendarType
//
// Every value came from an explicit locale name with the user override disabled, so neither this
// machine's display language nor its Regional Settings can appear in the data. Another Windows build
// answers differently in places, and that difference is the review surface when this is read again.

export interface NlsDateFormat {
  /** LOCALE_SSHORTDATE, a Windows date picture string (e.g. "dd.MM.yyyy"). */
  shortDate: string;
  /** LOCALE_SLONGDATE, the long date picture string with weekday and month names (e.g. "dddd, d. MMMM yyyy"). */
  longDate: string;
  /** LOCALE_STIMEFORMAT, a Windows time picture string (e.g. "HH:mm:ss"). */
  timeFormat: string;
  /** LOCALE_SSHORTTIME, the short time picture string (e.g. "HH:mm"). */
  shortTime: string;
  /** LOCALE_S1159 / LOCALE_S2359, the AM and PM designators ('' where the locale has none). */
  am: string;
  pm: string;
  /** LOCALE_ICALENDARTYPE: 1 = Gregorian, 6 = Hijri, 7 = Thai Buddhist, ... */
  calendarType: number;
}

const PACKED = `1025|dd/MM/yyyy|dd MMMM, yyyy|hh:mm:ss tt|hh:mm tt|ص|م|23
1026|d.M.yyyy 'г.'|dd MMMM yyyy 'г.'|H:mm:ss|H:mm|||1
1027|d/M/yyyy|dddd, d MMMM 'de' yyyy|H:mm:ss|H:mm|a.\u202fm.|p.\u202fm.|1
1028|yyyy/M/d|yyyy'年'M'月'd'日'|tt hh:mm:ss|tt hh:mm|上午|下午|1
1029|dd.MM.yyyy|dddd d. MMMM yyyy|H:mm:ss|H:mm|dop.|odp.|1
1030|dd-MM-yyyy|d. MMMM yyyy|HH:mm:ss|HH:mm|||1
1031|dd.MM.yyyy|dddd, d. MMMM yyyy|HH:mm:ss|HH:mm|||1
1032|d/M/yyyy|dddd, d MMMM yyyy|h:mm:ss tt|h:mm tt|πμ|μμ|1
1033|M/d/yyyy|dddd, MMMM d, yyyy|h:mm:ss tt|h:mm tt|AM|PM|1
1034|dd/MM/yyyy|dddd, dd' de 'MMMM' de 'yyyy|H:mm:ss|H:mm|||1
1035|d.M.yyyy|dddd d. MMMM yyyy|H.mm.ss|H.mm|ap.|ip.|1
1036|dd/MM/yyyy|dddd d MMMM yyyy|HH:mm:ss|HH:mm|||1
1037|dd/MM/yyyy|dddd dd MMMM yyyy|HH:mm:ss|HH:mm|AM|PM|1
1038|yyyy. MM. dd.|yyyy. MMMM d., dddd|H:mm:ss|H:mm|de.|du.|1
1039|d.M.yyyy|dddd, d. MMMM yyyy|HH:mm:ss|HH:mm|f.h.|e.h.|1
1040|dd/MM/yyyy|dddd d MMMM yyyy|HH:mm:ss|HH:mm|||1
1041|yyyy/MM/dd|yyyy'年'M'月'd'日'|H:mm:ss|H:mm|午前|午後|1
1042|yyyy-MM-dd|yyyy'년' M'월' d'일' dddd|tt h:mm:ss|tt h:mm|오전|오후|1
1043|d-M-yyyy|dddd d MMMM yyyy|HH:mm:ss|HH:mm|||1
1044|dd.MM.yyyy|dddd d. MMMM yyyy|HH:mm:ss|HH:mm|a.m.|p.m.|1
1045|d.MM.yyyy|dddd, d MMMM yyyy|HH:mm:ss|HH:mm|AM|PM|1
1046|dd/MM/yyyy|dddd, d' de 'MMMM' de 'yyyy|HH:mm:ss|HH:mm|||1
1047|dd-MM-yyyy|dddd, 'ils' d MMMM yyyy|HH:mm:ss|HH:mm|AM|PM|1
1048|dd.MM.yyyy|dddd, d MMMM yyyy|HH:mm:ss|HH:mm|a.m.|p.m.|1
1049|dd.MM.yyyy|d MMMM yyyy 'г.'|H:mm:ss|H:mm|||1
1050|d.M.yyyy.|d. MMMM yyyy.|H:mm:ss|H:mm|||1
1051|d. M. yyyy|dddd d. MMMM yyyy|H:mm:ss|H:mm|AM|PM|1
1052|d.M.yyyy|dddd, d MMMM yyyy|h:mm:ss\u202ftt|h:mm\u202ftt|p.d.|m.d.|1
1053|yyyy-MM-dd|'den 'd MMMM yyyy|HH:mm:ss|HH:mm|||1
1054|d/M/yyyy|'วัน'dddd'ที่' d MMMM yyyy|H:mm:ss|H:mm|AM|PM|7
1055|d.MM.yyyy|d MMMM yyyy dddd|HH:mm:ss|HH:mm|ÖÖ|ÖS|1
1056|dd/MM/yyyy|dd MMMM, yyyy|h:mm:ss tt|h:mm tt|AM|PM|1
1057|dd/MM/yyyy|dddd, dd MMMM yyyy|HH.mm.ss|HH.mm|AM|PM|1
1058|dd.MM.yyyy|d MMMM yyyy' р.'|H:mm:ss|H:mm|||1
1059|dd.MM.yy|d MMMM yyyy|HH:mm:ss|HH:mm|||1
1060|d. MM. yyyy|dddd, d. MMMM yyyy|HH:mm:ss|HH:mm|dop.|pop.|1
1061|dd.MM.yyyy|dddd, d. MMMM yyyy|HH:mm:ss|HH:mm|AM|PM|1
1062|dd.MM.yyyy|dddd, yyyy. 'gada' d. MMMM|HH:mm:ss|HH:mm|priekšp.|pēcp.|1
1063|yyyy-MM-dd|yyyy 'm'. MMMM d 'd'., dddd|HH:mm:ss|HH:mm|priešpiet|popiet|1
1065|dd/MM/yyyy|dddd, dd MMMM yyyy|hh:mm:ss tt|hh:mm tt|ق.ظ|ب.ظ|22
1066|dd/MM/yyyy|dd MMMM yyyy|h:mm:ss tt|h:mm tt|SA|CH|1
1067|dd.MM.yyyy|d MMMM, yyyy|HH:mm:ss|HH:mm|||1
1068|dd.MM.yyyy|d MMMM yyyy, dddd|HH:mm:ss|HH:mm|AM|PM|1
1069|yyyy/M/d|yyyy('e')'ko' MMMM'ren' d('a'), dddd|HH:mm:ss|HH:mm|AM|PM|1
1070|d.M.yyyy|dddd, d. MMMM yyyy|H:mm:ss|H:mm 'hodź'.|dop.|pop.|1
1071|d.M.yyyy|dddd, d MMMM yyyy|HH:mm:ss|HH:mm|претпл.|попл.|1
1072|yyyy-MM-dd|yyyy MMMM d, dddd|HH:mm:ss|HH:mm|AM|PM|1
1073|yyyy-MM-dd|yyyy MMMM d, dddd|HH:mm:ss|HH:mm|AM|PM|1
1074|yyyy-MM-dd|yyyy MMMM d, dddd|HH:mm:ss|HH:mm|AM|PM|1
1076|M/d/yyyy|dddd, MMMM d, yyyy|HH:mm:ss|HH:mm|AM|PM|1
1077|M/d/yyyy|dddd, MMMM d, yyyy|HH:mm:ss|HH:mm|AM|PM|1
1078|yyyy-MM-dd|dddd dd MMMM yyyy|HH:mm:ss|HH:mm|vm.|nm.|1
1080|dd.MM.yyyy|dddd, d. MMMM yyyy|HH:mm:ss|HH:mm|um fyr.|um sein.|1
1081|dd-MM-yyyy|dd MMMM yyyy|HH:mm:ss|HH:mm|पूर्वाह्न|अपराह्न|1
1082|dd/MM/yyyy|dddd, d 'ta'’ MMMM yyyy|HH:mm:ss|HH:mm|AM|PM|1
1085|dd/MM/yyyy|dddd, dטן MMMM yyyy|HH:mm:ss|HH:mm|פֿאַרמיטאָג|נאָכמיטאָג|1
1086|d/MM/yyyy|dddd, d MMMM yyyy|h:mm:ss\u202ftt|h:mm\u202ftt|PG|PTG|1
1089|dd/MM/yyyy|dddd, d MMMM yyyy|HH:mm:ss|HH:mm|AM|PM|1
1091|dd/MM/yyyy|dddd, d-MMMM, yyyy|HH:mm:ss|HH:mm|TO|TK|1
1092|dd.MM.yyyy|dd MMMM yyyy' ел'|HH:mm:ss|HH:mm|||1
1097|dd-MM-yyyy|dd MMMM yyyy|HH:mm:ss|HH:mm|காலை|மாலை|1
1102|dd-MM-yyyy|dd MMMM yyyy|HH:mm:ss|HH:mm|म.पू.|म.नं.|1
1103|d/M/yyyy|dddd, d MMMM yyyy|h:mm:ss tt|h:mm tt|AM|PM|1
2049|dd/MM/yyyy|dd MMMM, yyyy|hh:mm:ss tt|hh:mm tt|ص|م|1
2052|yyyy/M/d|yyyy'年'M'月'd'日'|H:mm:ss|H:mm|上午|下午|1
2055|dd.MM.yyyy|dddd, d. MMMM yyyy|HH:mm:ss|HH:mm|AM|PM|1
2057|dd/MM/yyyy|dd MMMM yyyy|HH:mm:ss|HH:mm|AM|PM|1
2058|dd/MM/yyyy|dddd, d' de 'MMMM' de 'yyyy|hh:mm:ss tt|hh:mm tt|a. m.|p. m.|1
2060|dd-MM-yy|dddd d MMMM yyyy|HH:mm:ss|HH:mm|||1
2064|dd.MM.yyyy|dddd, d MMMM yyyy|HH:mm:ss|HH:mm|AM|PM|1
2067|d/MM/yyyy|dddd d MMMM yyyy|H:mm:ss|H:mm|||1
2068|dd.MM.yyyy|dddd d. MMMM yyyy|HH:mm:ss|HH:mm|f.m.|e.m.|1
2070|dd/MM/yyyy|d' de 'MMMM' de 'yyyy|HH:mm:ss|HH:mm|||1
2072|dd.MM.yyyy|dddd, d MMMM yyyy|HH:mm:ss|HH:mm|a.m.|p.m.|1
2073|dd.MM.yyyy|dddd, d MMMM yyyy\u202f'г'.|HH:mm:ss|HH:mm|AM|PM|1
2074|d.M.yyyy.|d. MMMM yyyy.|H:mm:ss|H:mm|||1
2077|yyyy-MM-dd|dddd d MMMM yyyy|HH:mm:ss|HH:mm|fm|em|1
2092|dd.MM.yyyy|d MMMM yyyy|HH:mm:ss|H:mm|||1
2108|dd/MM/yyyy|dddd d MMMM yyyy|HH:mm:ss|HH:mm|r.n.|i.n.|1
2110|d/MM/yyyy|dd MMMM yyyy|h:mm:ss\u202ftt|h:mm\u202ftt|PG|PTG|1
2115|dd/MM/yyyy|dddd, dd MMMM, yyyy|HH:mm:ss|HH:mm|ТО|ТК|1
3073|dd/MM/yyyy|dd MMMM, yyyy|hh:mm:ss tt|hh:mm tt|ص|م|1
3076|d/M/yyyy|yyyy'年'M'月'd'日'|H:mm:ss|H:mm|上午|下午|1
3079|dd.MM.yyyy|dddd, d. MMMM yyyy|HH:mm:ss|HH:mm|AM|PM|1
3081|d/MM/yyyy|dddd, d MMMM yyyy|h:mm:ss tt|h:mm tt|AM|PM|1
3084|yyyy-MM-dd|d MMMM yyyy|HH:mm:ss|HH:mm|||1
3098|d.M.yyyy.|d. MMMM yyyy.|H:mm:ss|H:mm|||1
4097|dd/MM/yyyy|dd MMMM, yyyy|hh:mm:ss tt|hh:mm tt|ص|م|1
4100|d/M/yyyy|yyyy'年'M'月'd'日'|tt h:mm:ss|tt h:mm|上午|下午|1
4103|dd.MM.yyyy|dddd, d. MMMM yyyy|HH:mm:ss|HH:mm|AM|PM|1
4105|yyyy-MM-dd|MMMM d, yyyy|h:mm:ss tt|h:mm tt|AM|PM|1
4106|d/MM/yyyy|dddd, d 'de' MMMM 'de' yyyy|HH:mm:ss|HH:mm|a.\u202fm.|p.\u202fm.|1
4108|dd.MM.yyyy|dddd, d MMMM yyyy|HH:mm:ss|HH:mm|AM|PM|1
5121|dd-MM-yyyy|dd MMMM, yyyy|H:mm:ss|H:mm|ص|م|1
5124|d/M/yyyy|yyyy'年'M'月'd'日'|H:mm:ss|H:mm|上午|下午|1
5127|dd.MM.yyyy|dddd, d. MMMM yyyy|HH:mm:ss|HH:mm|AM|PM|1
5129|d/MM/yyyy|dddd, d MMMM yyyy|h:mm:ss tt|h:mm tt|am|pm|1
5130|d/M/yyyy|dddd, d 'de' MMMM 'de' yyyy|HH:mm:ss|HH:mm|a.\u202fm.|p.\u202fm.|1
5132|dd/MM/yyyy|dddd d MMMM yyyy|HH:mm:ss|HH:mm|AM|PM|1
6145|dd-MM-yyyy|dd MMMM, yyyy|H:mm:ss|H:mm|ص|م|1
6153|dd/MM/yyyy|dddd d MMMM yyyy|HH:mm:ss|HH:mm|am|pm|1
6154|MM/dd/yyyy|dddd, d 'de' MMMM 'de' yyyy|h:mm:ss\u202ftt|h:mm\u202ftt|a.\u202fm.|p.\u202fm.|1
7169|dd-MM-yyyy|dd MMMM, yyyy|H:mm:ss|H:mm|ص|م|1
7177|yyyy/MM/dd|dddd, dd MMMM yyyy|HH:mm:ss|HH:mm|am|pm|1
7178|d/M/yyyy|dddd, d 'de' MMMM 'de' yyyy|h:mm:ss\u202ftt|h:mm\u202ftt|a.\u202fm.|p.\u202fm.|1
8193|dd/MM/yyyy|dd MMMM, yyyy|hh:mm:ss tt|hh:mm tt|ص|م|1
8201|d/M/yyyy|dddd, d MMMM yyyy|h:mm:ss tt|h:mm tt|am|pm|1
8202|d/M/yyyy|dddd, d 'de' MMMM 'de' yyyy|h:mm:ss\u202ftt|h:mm\u202ftt|a.\u202fm.|p.\u202fm.|1
9217|dd/MM/yyyy|dd MMMM, yyyy|hh:mm:ss tt|hh:mm tt|ص|م|1
9225|dd/MM/yyyy|dddd, dd MMMM yyyy|HH:mm:ss|HH:mm|AM|PM|1
9226|d/MM/yyyy|dddd, d 'de' MMMM 'de' yyyy|h:mm:ss\u202ftt|h:mm\u202ftt|a.\u202fm.|p.\u202fm.|1
10241|dd/MM/yyyy|dd MMMM, yyyy|hh:mm:ss tt|hh:mm tt|ص|م|1
10249|dd/MM/yyyy|dddd, dd MMMM yyyy|HH:mm:ss|HH:mm|am|pm|1
10250|d/MM/yyyy|dddd, d 'de' MMMM 'de' yyyy|HH:mm:ss|HH:mm|a.\u202fm.|p.\u202fm.|1
11265|dd/MM/yyyy|dd MMMM, yyyy|hh:mm:ss tt|hh:mm tt|ص|م|1
11273|dd/MM/yyyy|dddd, d MMMM yyyy|h:mm:ss tt|h:mm tt|am|pm|1
11274|d/M/yyyy|dddd, d 'de' MMMM 'de' yyyy|HH:mm:ss|HH:mm|a.\u202fm.|p.\u202fm.|1
12289|dd/MM/yyyy|dd MMMM, yyyy|hh:mm:ss tt|hh:mm tt|ص|م|1
12298|d/M/yyyy|dddd, d 'de' MMMM 'de' yyyy|H:mm:ss|H:mm|a.\u202fm.|p.\u202fm.|1
13313|dd/MM/yyyy|dd MMMM, yyyy|hh:mm:ss tt|hh:mm tt|ص|م|1
13321|M/d/yyyy|dddd, MMMM d, yyyy|h:mm:ss tt|h:mm tt|AM|PM|1
13322|dd-MM-yyyy|dddd, d 'de' MMMM 'de' yyyy|H:mm:ss|H:mm|a.\u202fm.|p.\u202fm.|1
14337|dd/MM/yyyy|dd MMMM, yyyy|hh:mm:ss tt|hh:mm tt|ص|م|1
14346|d/M/yyyy|dddd, d 'de' MMMM 'de' yyyy|H:mm:ss|H:mm|a.\u202fm.|p.\u202fm.|1
15361|dd/MM/yyyy|dd MMMM, yyyy|hh:mm:ss tt|hh:mm tt|ص|م|1
15370|d/M/yyyy|dddd, d 'de' MMMM 'de' yyyy|HH:mm:ss|HH:mm|a.\u202fm.|p.\u202fm.|1
16385|dd/MM/yyyy|dd MMMM, yyyy|hh:mm:ss tt|hh:mm tt|ص|م|1
16394|d/M/yyyy|dddd, d 'de' MMMM 'de' yyyy|HH:mm:ss|HH:mm|a.\u202fm.|p.\u202fm.|1
17418|d/M/yyyy|dddd, d 'de' MMMM 'de' yyyy|HH:mm:ss|HH:mm|a.\u202fm.|p.\u202fm.|1
18442|d/M/yyyy|dddd dd 'de' MMMM 'de' yyyy|HH:mm:ss|HH:mm|a.\u202fm.|p.\u202fm.|1
19466|d/M/yyyy|dddd, d 'de' MMMM 'de' yyyy|HH:mm:ss|HH:mm|a.\u202fm.|p.\u202fm.|1
20490|MM/dd/yyyy|dddd, d 'de' MMMM 'de' yyyy|h:mm:ss\u202ftt|h:mm\u202ftt|a.\u202fm.|p.\u202fm.|1`;

export const NLS_DATE_FORMATS: ReadonlyMap<number, NlsDateFormat> = new Map(
  PACKED.split('\n').map(row => {
    const [lcid, shortDate, longDate, timeFormat, shortTime, am, pm, calendarType] = row.split('|');
    return [
      Number(lcid),
      { shortDate, longDate, timeFormat, shortTime, am, pm, calendarType: Number(calendarType) },
    ] as const;
  })
);

/** The NLS formats for an LCID, or undefined when the table has no row for it. */
export function nlsDateFormat(lcid: number): NlsDateFormat | undefined {
  return NLS_DATE_FORMATS.get(lcid);
}
