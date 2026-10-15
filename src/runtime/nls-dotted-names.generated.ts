// The month names Windows writes that end with a dot, per locale.
//
// Measured on Windows build 26200 25H2, 2026-10-15, by asking MonthName for every month, long and short, under
// every locale that answers. These are the names whose dot is part of the name rather than the dot of an
// abbreviation, and a fresh cscript reads one only where a space or the end of the text follows it: "15 jan.
// 2026" is January under 1038 while "15/jan./2026" and "15 jan., 2026" are type mismatches there. Thai is the
// one locale that reads its dotted names wherever they stand ("15/ม.ค./2026" is January), so its names are not
// used this way.
//
// 32 of the 132 locales write such a name. Editing this by hand makes it something other than what
// Windows answers.

/** The dotted month names Windows writes, in lower case, or an empty list for a locale that writes none. */
export const NLS_DOTTED_WRITTEN_NAMES: ReadonlyMap<number, readonly string[]> = new Map([
  [1027, ['abr.', 'ag.', 'des.', 'febr.', 'gen.', 'jul.', 'nov.', 'oct.', 'set.']],
  [1036, ['avr.', 'déc.', 'févr.', 'janv.', 'juil.', 'nov.', 'oct.', 'sept.']],
  [
    1038,
    [
      'aug.',
      'dec.',
      'febr.',
      'jan.',
      'júl.',
      'jún.',
      'máj.',
      'márc.',
      'nov.',
      'okt.',
      'szept.',
      'ápr.',
    ],
  ],
  [1039, ['apr.', 'des.', 'feb.', 'jan.', 'júl.', 'jún.', 'mar.', 'nóv.', 'okt.', 'sep.', 'ágú.']],
  [1047, ['avr.', 'dec.', 'fan.', 'favr.', 'nov.', 'oct.', 'schan.', 'sett.', 'zercl.']],
  [1048, ['apr.', 'aug.', 'dec.', 'feb.', 'ian.', 'iul.', 'iun.', 'mar.', 'nov.', 'oct.', 'sept.']],
  [
    1054,
    [
      'ก.ค.',
      'ก.พ.',
      'ก.ย.',
      'ต.ค.',
      'ธ.ค.',
      'พ.ค.',
      'พ.ย.',
      'ม.ค.',
      'มิ.ย.',
      'มี.ค.',
      'ส.ค.',
      'เม.ย.',
    ],
  ],
  [1060, ['apr.', 'avg.', 'dec.', 'feb.', 'jan.', 'jul.', 'jun.', 'mar.', 'nov.', 'okt.', 'sep.']],
  [1062, ['apr.', 'aug.', 'dec.', 'febr.', 'janv.', 'jūl.', 'jūn.', 'nov.', 'okt.', 'sept.']],
  [
    1063,
    [
      'bal.',
      'birž.',
      'geg.',
      'gruod.',
      'kov.',
      'lapkr.',
      'liep.',
      'rugp.',
      'rugs.',
      'saus.',
      'spal.',
      'vas.',
    ],
  ],
  [
    1069,
    [
      'abe.',
      'abu.',
      'api.',
      'aza.',
      'eka.',
      'ira.',
      'mai.',
      'mar.',
      'ots.',
      'urr.',
      'urt.',
      'uzt.',
    ],
  ],
  [
    1071,
    ['авг.', 'апр.', 'дек.', 'мар.', 'ноем.', 'окт.', 'септ.', 'фев.', 'јан.', 'јул.', 'јун.'],
  ],
  [1078, ['apr.', 'aug.', 'des.', 'feb.', 'jan.', 'jul.', 'jun.', 'mrt.', 'nov.', 'okt.', 'sep.']],
  [1092, ['авг.', 'апр.', 'гыйн.', 'дек.', 'мар.', 'нояб.', 'окт.', 'сен.', 'фев.']],
  [1102, ['ऑक्टो.', 'ऑग.', 'जाने.', 'डिसें.', 'नोव्हें.', 'फेब्रु.', 'सप्टें.']],
  [
    2058,
    [
      'abr.',
      'ago.',
      'dic.',
      'ene.',
      'feb.',
      'jul.',
      'jun.',
      'mar.',
      'may.',
      'nov.',
      'oct.',
      'sep.',
    ],
  ],
  [2060, ['avr.', 'déc.', 'févr.', 'janv.', 'juil.', 'nov.', 'oct.', 'sept.']],
  [2068, ['apr.', 'aug.', 'des.', 'feb.', 'jan.', 'nov.', 'okt.', 'sep.']],
  [2072, ['apr.', 'aug.', 'dec.', 'feb.', 'ian.', 'iul.', 'iun.', 'mar.', 'nov.', 'oct.', 'sept.']],
  [2073, ['авг.', 'апр.', 'дек.', 'нояб.', 'окт.', 'сент.', 'февр.', 'янв.']],
  [2074, ['apr.', 'avg.', 'dec.', 'feb.', 'jan.', 'nov.', 'okt.', 'sept.']],
  [2077, ['apr.', 'aug.', 'dec.', 'feb.', 'jan.', 'nov.', 'okt.', 'sep.']],
  [3084, ['avr.', 'déc.', 'févr.', 'janv.', 'juil.', 'nov.', 'oct.', 'sept.']],
  [3098, ['авг.', 'апр.', 'дец.', 'нов.', 'окт.', 'септ.', 'феб.', 'јан.']],
  [4108, ['avr.', 'déc.', 'févr.', 'janv.', 'juil.', 'nov.', 'oct.', 'sept.']],
  [5132, ['avr.', 'déc.', 'févr.', 'janv.', 'juil.', 'nov.', 'oct.', 'sept.']],
  [
    8202,
    [
      'abr.',
      'ago.',
      'dic.',
      'ene.',
      'feb.',
      'jul.',
      'jun.',
      'mar.',
      'may.',
      'nov.',
      'oct.',
      'sept.',
    ],
  ],
  [
    9226,
    [
      'abr.',
      'ago.',
      'dic.',
      'ene.',
      'feb.',
      'jul.',
      'jun.',
      'mar.',
      'may.',
      'nov.',
      'oct.',
      'sept.',
    ],
  ],
  [
    10250,
    [
      'abr.',
      'ago.',
      'dic.',
      'ene.',
      'feb.',
      'jul.',
      'jun.',
      'mar.',
      'may.',
      'nov.',
      'oct.',
      'set.',
    ],
  ],
  [
    13322,
    [
      'abr.',
      'ago.',
      'dic.',
      'ene.',
      'feb.',
      'jul.',
      'jun.',
      'mar.',
      'may.',
      'nov.',
      'oct.',
      'sept.',
    ],
  ],
  [
    14346,
    [
      'abr.',
      'ago.',
      'dic.',
      'ene.',
      'feb.',
      'jul.',
      'jun.',
      'mar.',
      'may.',
      'nov.',
      'oct.',
      'set.',
    ],
  ],
  [
    15370,
    [
      'abr.',
      'ago.',
      'dic.',
      'ene.',
      'feb.',
      'jul.',
      'jun.',
      'mar.',
      'may.',
      'nov.',
      'oct.',
      'sept.',
    ],
  ],
]);

/** The dotted names of one locale, for the reader that asks whether a dot closes a name. */
export function dottedWrittenNames(lcid: number): ReadonlySet<string> {
  const cached = dottedNameCache.get(lcid);
  if (cached) return cached;
  const names = new Set(NLS_DOTTED_WRITTEN_NAMES.get(lcid) ?? []);
  dottedNameCache.set(lcid, names);
  return names;
}

const dottedNameCache: Map<number, ReadonlySet<string>> = new Map();
