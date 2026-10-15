import type { VbValue } from '../runtime/index.ts';
import { LCID_TO_LOCALE } from './locale-data.generated.ts';

const localeToLcid: Map<string, number> = new Map([
  ['af', 1078],
  ['sq', 1052],
  ['ar-ae', 14337],
  ['ar-bh', 15361],
  ['ar-dz', 5121],
  ['ar-eg', 3073],
  ['ar-iq', 2049],
  ['ar-jo', 11265],
  ['ar-kw', 13313],
  ['ar-lb', 12289],
  ['ar-ly', 4097],
  ['ar-ma', 6145],
  ['ar-om', 8193],
  ['ar-qa', 16385],
  ['ar-sa', 1025],
  ['ar-sy', 10241],
  ['ar-tn', 7169],
  ['ar-ye', 9217],
  ['hy', 1067],
  ['az', 1068],
  ['az-Latn', 1068],
  ['az-Cyrl', 2092],
  ['eu', 1069],
  ['be', 1059],
  ['bg', 1026],
  ['ca', 1027],
  ['zh-cn', 2052],
  ['zh-hk', 3076],
  ['zh-mo', 5124],
  ['zh-sg', 4100],
  ['zh-tw', 1028],
  ['zh', 2052],
  ['hr', 1050],
  ['cs', 1029],
  ['da', 1030],
  ['nl-nl', 1043],
  ['nl-be', 2067],
  ['nl', 1043],
  ['en-au', 3081],
  ['en-bz', 10249],
  ['en-ca', 4105],
  ['en-cb', 9225],
  ['en-ie', 6153],
  ['en-jm', 8201],
  ['en-nz', 5129],
  ['en-ph', 13321],
  ['en-za', 7177],
  ['en-tt', 11273],
  ['en-gb', 2057],
  ['en-us', 1033],
  ['en', 1033],
  ['et', 1061],
  ['fa', 1065],
  ['fi', 1035],
  ['fo', 1080],
  ['fr-fr', 1036],
  ['fr-be', 2060],
  ['fr-ca', 3084],
  ['fr-lu', 5132],
  ['fr-ch', 4108],
  ['fr', 1036],
  ['gd-ie', 2108],
  ['gd', 1084],
  ['de-de', 1031],
  ['de-at', 3079],
  ['de-li', 5127],
  ['de-lu', 4103],
  ['de-ch', 2055],
  ['de', 1031],
  ['el', 1032],
  ['he', 1037],
  ['hi', 1081],
  ['hu', 1038],
  ['is', 1039],
  ['id', 1057],
  ['it-it', 1040],
  ['it-ch', 2064],
  ['it', 1040],
  ['ja', 1041],
  ['ko', 1042],
  ['lv', 1062],
  ['lt', 1063],
  ['mk', 1071],
  ['ms-my', 1086],
  ['ms-bn', 2110],
  ['ms', 1086],
  ['mt', 1082],
  ['mr', 1102],
  ['no', 1044],
  ['nb', 1044],
  ['nn', 2068],
  ['pl', 1045],
  ['pt-pt', 2070],
  ['pt-br', 1046],
  ['pt', 1046],
  ['rm', 1047],
  ['ro', 1048],
  ['ro-md', 2072],
  ['ru', 1049],
  ['ru-md', 2073],
  ['sa', 1103],
  ['sr-Cyrl', 3098],
  ['sr-Latn', 2074],
  ['sr', 3098],
  ['tn', 1074],
  ['sl', 1060],
  ['sk', 1051],
  ['sb', 1070],
  ['es-es', 1034],
  ['es-ar', 11274],
  ['es-bo', 16394],
  ['es-cl', 13322],
  ['es-co', 9226],
  ['es-cr', 5130],
  ['es-do', 7178],
  ['es-ec', 12298],
  ['es-gt', 4106],
  ['es-hn', 18442],
  ['es-mx', 2058],
  ['es-ni', 19466],
  ['es-pa', 6154],
  ['es-pe', 10250],
  ['es-pr', 20490],
  ['es-py', 15370],
  ['es-sv', 17418],
  ['es-uy', 14346],
  ['es-ve', 8202],
  ['es', 1034],
  ['sx', 1072],
  ['sw', 1089],
  ['sv-se', 1053],
  ['sv-fi', 2077],
  ['sv', 1053],
  ['ta', 1097],
  ['tt', 1092],
  ['th', 1054],
  ['tr', 1055],
  ['ts', 1073],
  ['uk', 1058],
  ['ur', 1056],
  ['uz-Cyrl', 2115],
  ['uz-Latn', 1091],
  ['uz', 1091],
  ['vi', 1066],
  ['xh', 1076],
  ['yi', 1085],
  ['zu', 1077],
]);

// Windows' own name for each LCID, read from LCIDToLocaleName and kept only where ICU resolves it exactly.
// Inverting the hand-written map instead gave several LCIDs the wrong locale: the names for 1070 and
// 1072 ("sb", "sx") resolve to nothing, so both were read as the host locale, and 2108 was read as
// Scottish Gaelic rather than Irish.
const lcidToLocale: Map<number, string> = new Map(LCID_TO_LOCALE);

// Every generated name has to resolve back to its LCID, or GetLocale() after SetLocale(n) would fall
// back to the default. Keys that already exist stay first, so today's spelling keeps winning.
LCID_TO_LOCALE.forEach((locale, lcid) => {
  if (!localeToLcid.has(locale)) {
    localeToLcid.set(locale, lcid);
  }
});

/**
 * The locale of the engine that is currently running.
 *
 * A script host runs one script at a time, so the engine installs its own state around each call. Two
 * engines in one process can then hold different locales, and one engine's SetLocale is no longer
 * visible to another - which it was while this was a single module-level variable, shared by every
 * engine in the process.
 */
export interface LocaleState {
  locale: string | null;
}

let ambientLocaleState: LocaleState | null = null;
/** Used only when a builtin is called with no engine running; no script can reach this. */
let detachedLocale: string | null = null;

/** Runs `body` with `state` installed as the locale of the engine that is running. */
export function withLocaleState<T>(state: LocaleState, body: () => T): T {
  const previous = ambientLocaleState;
  ambientLocaleState = state;
  try {
    return body();
  } finally {
    ambientLocaleState = previous;
  }
}

function currentLocale(): string | null {
  return ambientLocaleState === null ? detachedLocale : ambientLocaleState.locale;
}

function rememberLocale(locale: string): void {
  if (ambientLocaleState === null) {
    detachedLocale = locale;
  } else {
    ambientLocaleState.locale = locale;
  }
}

function normalizeLocale(locale: string): string {
  return locale.toLowerCase().replace(/_/g, '-');
}

function findLocaleKey(normalized: string): string | undefined {
  for (const key of localeToLcid.keys()) {
    if (key.toLowerCase() === normalized) {
      return key;
    }
  }
  return undefined;
}

function getBrowserLocale(): string {
  const chosen = currentLocale();
  if (chosen) {
    return chosen;
  }

  const browserLocale = Intl.DateTimeFormat().resolvedOptions().locale || 'en-US';
  const normalized = normalizeLocale(browserLocale);

  const exactKey = findLocaleKey(normalized);
  if (exactKey) {
    return exactKey;
  }

  const baseLocale = normalized.split('-')[0] ?? 'en';
  const baseKey = findLocaleKey(baseLocale);
  if (baseKey) {
    return baseKey;
  }

  return 'en';
}

/** The LCID of the locale the engine is running under. */
export function currentLcid(): number {
  return localeToLcid.get(getBrowserLocale()) ?? 1033;
}

export function getLocale(): VbValue {
  return { type: 'Long', value: currentLcid() };
}

export function setLocale(lcid: VbValue): VbValue {
  const lcidValue = typeof lcid.value === 'number' ? lcid.value : parseInt(String(lcid.value), 10);

  if (isNaN(lcidValue)) {
    const localeStr = String(lcid.value).toLowerCase();
    const key = findLocaleKey(localeStr);
    if (key) {
      rememberLocale(key);
      return { type: 'Long', value: localeToLcid.get(key) ?? 1033 };
    }
    return getLocale();
  }

  const locale = lcidToLocale.get(lcidValue);
  if (locale) {
    rememberLocale(locale);
    return { type: 'Long', value: lcidValue };
  }

  return getLocale();
}

export function getCurrentLocaleTag(): string {
  const locale = getBrowserLocale();
  const parts = locale.replace(/-/g, '_').split('_');
  return parts[0] + (locale.includes('-') ? '-' + (parts[1]?.toUpperCase() ?? '') : '');
}

export function getCurrentBCP47Locale(): string {
  return getBrowserLocale();
}

// The currency a locale writes with is not decided here. It comes from the locale's own row in the Windows
// NLS number table, which is where vbscript.dll reads it, along with the two sets of marks and the shape of
// a negative amount. A hand-written map used to live here and gave the United States dollar to every locale
// it did not list, which included two of the German family: de-LU writes EUR and de-LI writes CHF.

export const localeFunctions = {
  GetLocale: getLocale,
  SetLocale: setLocale,
};
