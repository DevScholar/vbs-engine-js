// Windows NLS number and money formats. This file is data read from Windows, not a build artifact, and
// editing it by hand makes it something other than what Windows answers.
//
// Read with GetLocaleInfoEx(locale, LOCALE_NOUSEROVERRIDE):
//   LOCALE_SDECIMAL, LOCALE_STHOUSAND, LOCALE_SGROUPING, LOCALE_IDIGITS, LOCALE_ILZERO,
//   LOCALE_INEGNUMBER, LOCALE_SNEGATIVESIGN, LOCALE_SPOSITIVESIGN, LOCALE_SCURRENCY,
//   LOCALE_SINTLSYMBOL, LOCALE_ICURRDIGITS, LOCALE_ICURRENCY, LOCALE_INEGCURR,
//   LOCALE_SMONDECIMALSEP, LOCALE_SMONTHOUSANDSEP, LOCALE_SMONGROUPING, LOCALE_SNATIVEDIGITS
//
// Read on Windows build 26200 25H2, 2026-10-15
// Rows: lcid|decimal|thousand|grouping|digits|leadingZero|negative|negativeSign|positiveSign|symbol|code|currencyDigits|currencyPosition|negativeCurrency|monDecimal|monThousand|monGrouping|nativeDigits
//
// Money has its own two marks in Windows (LOCALE_SMONDECIMALSEP, LOCALE_SMONTHOUSANDSEP), and they are
// not always the ones LOCALE_SDECIMAL and LOCALE_STHOUSAND hold: de-AT groups a plain number with U+00A0
// and money with a full stop. A locale's row therefore carries both sets.
//
// Every value came from an explicit locale name with the user override disabled, so neither this
// machine's display language nor its Regional Settings can appear in the data. Another Windows build
// answers differently in places, and that difference is the review surface when this is read again.

/** One locale's Windows number format. Every field is data read from Windows, never inferred. */
export interface NlsNumberFormat {
  /** LOCALE_SDECIMAL: the mark between a whole number and its fraction. */
  decimal: string;
  /** LOCALE_STHOUSAND: the mark between groups in a plain number. */
  thousand: string;
  /** LOCALE_SGROUPING: group sizes from the right, e.g. "3;0" (every three, repeating). */
  grouping: string;
  /** LOCALE_IDIGITS: the fraction count FormatNumber uses when the caller does not name one. */
  digits: number;
  /** LOCALE_ILZERO: 1 writes a zero before a fraction ("0,5"), 0 leaves it out (",5"). */
  leadingZero: number;
  /** LOCALE_INEGNUMBER: 0 "(1)", 1 "-1", 2 "- 1", 3 "1-", 4 "1 -". */
  negativeNumber: number;
  /** LOCALE_SNEGATIVESIGN. */
  negativeSign: string;
  /** LOCALE_SPOSITIVESIGN. */
  positiveSign: string;
  /** LOCALE_SCURRENCY: the symbol FormatCurrency writes. */
  currencySymbol: string;
  /** LOCALE_SINTLSYMBOL: the ISO code, e.g. "EUR", as Windows reports it. */
  currencyCode: string;
  /** LOCALE_ICURRDIGITS: the fraction count the currency itself uses. */
  currencyDigits: number;
  /** LOCALE_ICURRENCY: 0 symbol first, 1 symbol last, 2 symbol first with a space, 3 symbol last with a space. */
  currencyPosition: number;
  /** LOCALE_INEGCURR: which of the sixteen ways a negative amount is written. */
  negativeCurrency: number;
  /** LOCALE_SMONDECIMALSEP: the fraction mark money uses. */
  monDecimal: string;
  /** LOCALE_SMONTHOUSANDSEP: the group mark money uses. */
  monThousand: string;
  /** LOCALE_SMONGROUPING: the group sizes money uses. */
  monGrouping: string;
  /** LOCALE_SNATIVEDIGITS: the locale's own digits, ten characters, ASCII for most locales. */
  nativeDigits: string;
}

const PACKED = `1025|.|,|3;0|2|1|1|-||\u20c1|SAR|2|2|12|.|,|3;0|\u0660\u0661\u0662\u0663\u0664\u0665\u0666\u0667\u0668\u0669
1026|,|\u00a0|3;0|2|1|1|-||\u043b\u0432.|BGN|2|3|8|,|\u00a0|3;0|0123456789
1027|,|.|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|.|3;0|0123456789
1028|.|,|3;0|2|1|1|-||NT\u0024|TWD|2|0|1|.|,|3;0|0123456789
1029|,|\u00a0|3;0|2|1|1|-||K\u010d|CZK|2|3|8|,|\u00a0|3;0|0123456789
1030|,|.|3;0|2|1|1|-||kr.|DKK|2|3|8|,|.|3;0|0123456789
1031|,|.|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|.|3;0|0123456789
1032|,|.|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|.|3;0|0123456789
1033|.|,|3;0|2|1|1|-||\u0024|USD|2|0|0|.|,|3;0|0123456789
1034|,|.|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|.|3;0|0123456789
1035|,|\u00a0|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|\u00a0|3;0|0123456789
1036|,|\u202f|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|\u202f|3;0|0123456789
1037|.|,|3;0|2|1|1|-||\u20aa|ILS|2|2|2|.|,|3;0|0123456789
1038|,|\u00a0|3;0|2|1|1|-||Ft|HUF|2|3|8|,|\u00a0|3;0|0123456789
1039|,|.|3;0|2|1|1|-||kr.|ISK|0|3|8|,|.|3;0|0123456789
1040|,|.|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|.|3;0|0123456789
1041|.|,|3;0|2|1|1|-||\u00a5|JPY|0|0|1|.|,|3;0|0123456789
1042|.|,|3;0|2|1|1|-||\u20a9|KRW|0|0|1|.|,|3;0|0123456789
1043|,|.|3;0|2|1|1|-||\u20ac|EUR|2|2|12|,|.|3;0|0123456789
1044|,|\u00a0|3;0|2|1|1|-||kr|NOK|2|2|12|,|\u00a0|3;0|0123456789
1045|,|\u00a0|3;0|2|1|1|-||z\u0142|PLN|2|3|8|,|\u00a0|3;0|0123456789
1046|,|.|3;0|2|1|1|-||R\u0024|BRL|2|2|9|,|.|3;0|0123456789
1047|.|\u2019|3;0|2|1|1|-||CHF|CHF|2|3|8|.|\u2019|3;0|0123456789
1048|,|.|3;0|2|1|1|-||lei|RON|2|3|8|,|.|3;0|0123456789
1049|,|\u00a0|3;0|2|1|1|-||\u20bd|RUB|2|3|8|,|\u00a0|3;0|0123456789
1050|,|.|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|.|3;0|0123456789
1051|,|\u00a0|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|\u00a0|3;0|0123456789
1052|,|\u00a0|3;0|2|1|1|-||Lek\u00eb|ALL|0|3|8|,|\u00a0|3;0|0123456789
1053|,|\u00a0|3;0|2|1|1|-||kr|SEK|2|3|8|,|\u00a0|3;0|0123456789
1054|.|,|3;0|2|1|1|-||\u0e3f|THB|2|0|1|.|,|3;0|\u0e50\u0e51\u0e52\u0e53\u0e54\u0e55\u0e56\u0e57\u0e58\u0e59
1055|,|.|3;0|2|1|1|-||\u20ba|TRY|2|0|1|,|.|3;0|0123456789
1056|.|,|3;0|2|1|1|-||Rs|PKR|2|0|1|.|,|3;0|\u06f0\u06f1\u06f2\u06f3\u06f4\u06f5\u06f6\u06f7\u06f8\u06f9
1057|,|.|3;0|2|1|1|-||Rp|IDR|2|0|1|,|.|3;0|0123456789
1058|,|\u00a0|3;0|2|1|1|-||\u20b4|UAH|2|3|8|,|\u00a0|3;0|0123456789
1059|,|\u00a0|3;0|2|1|1|-||Br|BYN|2|3|8|,|\u00a0|3;0|0123456789
1060|,|.|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|.|3;0|0123456789
1061|,|\u00a0|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|\u00a0|3;0|0123456789
1062|,|\u00a0|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|\u00a0|3;0|0123456789
1063|,|\u00a0|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|\u00a0|3;0|0123456789
1065|/|,|3;0|2|1|3|-||\u0631\u064a\u0627\u0644|IRR|2|1|6|/|,|3;0|\u06f0\u06f1\u06f2\u06f3\u06f4\u06f5\u06f6\u06f7\u06f8\u06f9
1066|,|.|3;0|2|1|1|-||\u20ab|VND|0|3|8|,|.|3;0|0123456789
1067|.|,|3;0|2|1|1|-||\u058f|AMD|2|3|8|.|,|3;0|0123456789
1068|,|.|3;0|2|1|1|-||\u20bc|AZN|2|3|8|,|.|3;0|0123456789
1069|,|.|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|.|3;0|0123456789
1070|,|.|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|.|3;0|0123456789
1071|,|.|3;0|2|1|1|-||\u0434\u0435\u043d.|MKD|2|3|8|,|.|3;0|0123456789
1072|,|\u00a0|3;0|2|1|1|-||R|ZAR|2|0|1|,|\u00a0|3;0|0123456789
1073|,|\u00a0|3;0|2|1|1|-||R|ZAR|2|2|9|,|\u00a0|3;0|0123456789
1074|.|\u00a0|3;0|2|1|1|-||R|ZAR|2|0|1|.|\u00a0|3;0|0123456789
1076|.|\u00a0|3;0|2|1|1|-||R|ZAR|2|0|1|.|\u00a0|3;0|0123456789
1077|.|,|3;0|2|1|1|-||R|ZAR|2|0|1|.|,|3;0|0123456789
1078|,|\u00a0|3;0|2|1|1|-||R|ZAR|2|0|1|,|\u00a0|3;0|0123456789
1080|,|.|3;0|2|1|1|-||kr|DKK|2|3|8|,|.|3;0|0123456789
1081|.|,|3;2;0|2|1|1|-||\u20b9|INR|2|0|12|.|,|3;2;0|\u0966\u0967\u0968\u0969\u096a\u096b\u096c\u096d\u096e\u096f
1082|.|,|3;0|2|1|1|-||\u20ac|EUR|2|0|1|.|,|3;0|0123456789
1085|.|,|3;0|2|1|1|-||XDR|XDR|2|2|9|.|,|3;0|0123456789
1086|.|,|3;0|2|1|1|-||RM|MYR|2|0|1|.|,|3;0|0123456789
1089|.|,|3;0|2|1|1|-||Ksh|KES|2|0|1|.|,|3;0|0123456789
1091|,|\u00a0|3;0|2|1|1|-||so\u02bbm|UZS|2|3|8|,|\u00a0|3;0|0123456789
1092|,|\u00a0|3;0|2|1|1|-||\u20bd|RUB|2|3|8|,|\u00a0|3;0|0123456789
1097|.|,|3;2;0|2|1|1|-||\u20b9|INR|2|2|12|.|,|3;2;0|\u0be6\u0be7\u0be8\u0be9\u0bea\u0beb\u0bec\u0bed\u0bee\u0bef
1102|.|,|3;2;0|2|1|1|-||\u20b9|INR|2|0|12|\`|,|3;0|\u0966\u0967\u0968\u0969\u096a\u096b\u096c\u096d\u096e\u096f
1103|.|,|3;2;0|2|1|1|-||\u20b9|INR|2|0|1|.|,|3;2;0|\u0966\u0967\u0968\u0969\u096a\u096b\u096c\u096d\u096e\u096f
2049|.|,|3;0|2|1|1|-||\u062f.\u0639.\u200f|IQD|2|3|8|.|,|3;0|\u0660\u0661\u0662\u0663\u0664\u0665\u0666\u0667\u0668\u0669
2052|.|,|3;0|2|1|1|-||\u00a5|CNY|2|0|2|.|,|3;0|0123456789
2055|.|\u2019|3;0|2|1|1|-||CHF|CHF|2|2|2|.|\u2019|3;0|0123456789
2057|.|,|3;0|2|1|1|-||\u00a3|GBP|2|0|1|.|,|3;0|0123456789
2058|.|,|3;0|2|1|1|-||\u0024|MXN|2|0|1|.|,|3;0|0123456789
2060|,|.|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|.|3;0|0123456789
2064|.|\u2019|3;0|2|1|1|-||CHF|CHF|2|2|2|.|\u2019|3;0|0123456789
2067|,|.|3;0|2|1|1|-||\u20ac|EUR|2|2|12|,|.|3;0|0123456789
2068|,|\u00a0|3;0|2|1|1|-||kr|NOK|2|3|8|,|\u00a0|3;0|0123456789
2070|,|\u00a0|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|\u00a0|3;0|0123456789
2072|,|.|3;0|2|1|1|-||L|MDL|2|3|8|,|.|3;0|0123456789
2073|,|\u00a0|3;0|2|1|1|-||L|MDL|2|3|8|,|\u00a0|3;0|0123456789
2074|,|.|3;0|2|1|1|-||din.|RSD|2|3|8|,|.|3;0|0123456789
2077|,|\u00a0|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|\u00a0|3;0|0123456789
2092|,|\u00a0|3;0|2|1|1|-||\u20bc|AZN|2|3|8|,|\u00a0|3;0|0123456789
2108|.|,|3;0|2|1|1|-||\u20ac|EUR|2|0|1|.|,|3;0|0123456789
2110|,|.|3;0|2|1|1|-||\u0024|BND|2|2|9|,|.|3;0|0123456789
2115|,|\u00a0|3;0|2|1|1|-||\u0441\u045e\u043c|UZS|2|3|8|,|\u00a0|3;0|0123456789
3073|.|,|3;0|2|1|1|-||\u062c.\u0645.\u200f|EGP|2|3|8|.|,|3;0|\u0660\u0661\u0662\u0663\u0664\u0665\u0666\u0667\u0668\u0669
3076|.|,|3;0|2|1|1|-||HK\u0024|HKD|2|0|0|.|,|3;0|0123456789
3079|,|\u00a0|3;0|2|1|1|-||\u20ac|EUR|2|2|9|,|.|3;0|0123456789
3081|.|,|3;0|2|1|1|-||\u0024|AUD|2|0|1|.|,|3;0|0123456789
3084|,|\u00a0|3;0|2|1|1|-||\u0024|CAD|2|3|15|,|\u00a0|3;0|0123456789
3098|,|.|3;0|2|1|1|-||\u0434\u0438\u043d.|RSD|2|3|8|,|.|3;0|0123456789
4097|.|,|3;0|2|1|1|-||\u062f.\u0644.\u200f|LYD|3|3|8|.|,|3;0|0123456789
4100|.|,|3;0|2|1|1|-||\u0024|SGD|2|0|1|.|,|3;0|0123456789
4103|,|.|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|.|3;0|0123456789
4105|.|,|3;0|2|1|1|-||\u0024|CAD|2|0|1|.|,|3;0|0123456789
4106|.|,|3;0|2|1|1|-||Q|GTQ|2|0|1|.|,|3;0|0123456789
4108|,|\u202f|3;0|2|1|1|-||CHF|CHF|2|3|8|.|\u202f|3;0|0123456789
5121|.|,|3;0|2|1|1|-||\u062f.\u062c.\u200f|DZD|2|3|8|.|,|3;0|0123456789
5124|.|,|3;0|2|1|1|-||MOP|MOP|2|0|1|.|,|3;0|0123456789
5127|.|\u2019|3;0|2|1|1|-||CHF|CHF|2|2|9|.|\u2019|3;0|0123456789
5129|.|,|3;0|2|1|1|-||\u0024|NZD|2|0|1|.|,|3;0|0123456789
5130|,|\u00a0|3;0|2|1|1|-||\u20a1|CRC|2|0|1|,|\u00a0|3;0|0123456789
5132|,|.|3;0|2|1|1|-||\u20ac|EUR|2|3|8|,|.|3;0|0123456789
6145|.|,|3;0|2|1|1|-||\u062f.\u0645.\u200f|MAD|2|3|8|.|,|3;0|0123456789
6153|.|,|3;0|2|1|1|-||\u20ac|EUR|2|0|1|.|,|3;0|0123456789
6154|.|,|3;0|2|1|1|-||B/.|PAB|2|0|1|.|,|3;0|0123456789
7169|.|,|3;0|2|1|1|-||\u062f.\u062a.\u200f|TND|3|3|8|.|,|3;0|0123456789
7177|,|\u00a0|3;0|2|1|1|-||R|ZAR|2|0|1|,|\u00a0|3;0|0123456789
7178|.|,|3;0|2|1|1|-||\u0024|DOP|2|0|0|.|,|3;0|0123456789
8193|.|,|3;0|2|1|1|-||\u0631.\u0639.\u200f|OMR|3|3|8|.|,|3;0|\u0660\u0661\u0662\u0663\u0664\u0665\u0666\u0667\u0668\u0669
8201|.|,|3;0|2|1|1|-||\u0024|JMD|2|0|1|.|,|3;0|0123456789
8202|,|.|3;0|2|1|1|-||Bs.S|VES|2|0|2|,|.|3;0|0123456789
9217|.|,|3;0|2|1|1|-||\u0631.\u064a.\u200f|YER|2|3|8|.|,|3;0|\u0660\u0661\u0662\u0663\u0664\u0665\u0666\u0667\u0668\u0669
9225|.|,|3;0|2|1|1|-||EC\u0024|XCD|2|0|1|.|,|3;0|0123456789
9226|,|.|3;0|2|1|1|-||\u0024|COP|2|2|9|,|.|3;0|0123456789
10241|.|,|3;0|2|1|1|-||\u0644.\u0633.\u200f|SYP|2|3|8|.|,|3;0|\u0660\u0661\u0662\u0663\u0664\u0665\u0666\u0667\u0668\u0669
10249|.|,|3;0|2|1|1|-||\u0024|BZD|2|0|1|.|,|3;0|0123456789
10250|.|,|3;0|2|1|1|-||S/|PEN|2|2|9|.|,|3;0|0123456789
11265|.|,|3;0|2|1|1|-||\u062f.\u0627.\u200f|JOD|3|3|8|.|,|3;0|\u0660\u0661\u0662\u0663\u0664\u0665\u0666\u0667\u0668\u0669
11273|.|,|3;0|2|1|1|-||\u0024|TTD|2|0|1|.|,|3;0|0123456789
11274|,|.|3;0|2|1|1|-||\u0024|ARS|2|2|9|,|.|3;0|0123456789
12289|.|,|3;0|2|1|1|-||\u0644.\u0644.\u200f|LBP|2|3|8|.|,|3;0|\u0660\u0661\u0662\u0663\u0664\u0665\u0666\u0667\u0668\u0669
12298|,|.|3;0|2|1|1|-||\u0024|USD|2|0|2|,|.|3;0|0123456789
13313|.|,|3;0|2|1|1|-||\u062f.\u0643.\u200f|KWD|3|3|8|.|,|3;0|\u0660\u0661\u0662\u0663\u0664\u0665\u0666\u0667\u0668\u0669
13321|.|,|3;0|2|1|1|-||\u20b1|PHP|2|0|1|.|,|3;0|0123456789
13322|,|.|3;0|2|1|1|-||\u0024|CLP|0|0|2|,|.|3;0|0123456789
14337|.|,|3;0|2|1|1|-||\u062f.\u0625.\u200f|AED|2|3|8|.|,|3;0|\u0660\u0661\u0662\u0663\u0664\u0665\u0666\u0667\u0668\u0669
14346|,|.|3;0|2|1|1|-||\u0024|UYU|2|2|9|,|.|3;0|0123456789
15361|.|,|3;0|2|1|1|-||\u062f.\u0628.\u200f|BHD|3|3|8|.|,|3;0|\u0660\u0661\u0662\u0663\u0664\u0665\u0666\u0667\u0668\u0669
15370|,|.|3;0|2|1|1|-||\u20b2|PYG|0|2|12|,|.|3;0|0123456789
16385|.|,|3;0|2|1|1|-||\u0631.\u0642.\u200f|QAR|2|3|8|.|,|3;0|\u0660\u0661\u0662\u0663\u0664\u0665\u0666\u0667\u0668\u0669
16394|,|.|3;0|2|1|1|-||Bs|BOB|2|0|1|,|.|3;0|0123456789
17418|.|,|3;0|2|1|1|-||\u0024|USD|2|0|1|.|,|3;0|0123456789
18442|.|,|3;0|2|1|1|-||L|HNL|2|0|1|.|,|3;0|0123456789
19466|.|,|3;0|2|1|1|-||C\u0024|NIO|2|0|1|.|,|3;0|0123456789
20490|.|,|3;0|2|1|1|-||\u0024|USD|2|0|1|.|,|3;0|0123456789`;

const FIELDS = PACKED.split('\n').map(row => row.split('|'));

/** The Windows number format for an LCID, or undefined when the table has no row for it. */
export const NLS_NUMBER_FORMATS: ReadonlyMap<number, NlsNumberFormat> = new Map(
  FIELDS.map(row => [
    Number(row[0]),
    {
      decimal: row[1],
      thousand: row[2],
      grouping: row[3],
      digits: Number(row[4]),
      leadingZero: Number(row[5]),
      negativeNumber: Number(row[6]),
      negativeSign: row[7],
      positiveSign: row[8],
      currencySymbol: row[9],
      currencyCode: row[10],
      currencyDigits: Number(row[11]),
      currencyPosition: Number(row[12]),
      negativeCurrency: Number(row[13]),
      monDecimal: row[14],
      monThousand: row[15],
      monGrouping: row[16],
      nativeDigits: row[17],
    },
  ])
);

/** The Windows number format for an LCID, or undefined when the table has no row for it. */
export function nlsNumberFormat(lcid: number): NlsNumberFormat | undefined {
  return NLS_NUMBER_FORMATS.get(lcid);
}
