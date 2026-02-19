import type { VbValue } from '../runtime/index.ts';

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

const lcidToLocale: Map<number, string> = new Map();
localeToLcid.forEach((lcid, locale) => {
  lcidToLocale.set(lcid, locale);
});

let currentLocale: string | null = null;

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
  if (currentLocale) {
    return currentLocale;
  }

  const browserLocale = Intl.DateTimeFormat().resolvedOptions().locale || 'en-US';
  const normalized = normalizeLocale(browserLocale);

  const exactKey = findLocaleKey(normalized);
  if (exactKey) {
    return exactKey;
  }

  const baseLocale = normalized.split('-')[0];
  const baseKey = findLocaleKey(baseLocale);
  if (baseKey) {
    return baseKey;
  }

  return 'en';
}

export function getLocale(): VbValue {
  const locale = getBrowserLocale();
  const lcid = localeToLcid.get(locale) ?? 1033;
  return { type: 'Long', value: lcid };
}

export function setLocale(lcid: VbValue): VbValue {
  const lcidValue = typeof lcid.value === 'number' ? lcid.value : parseInt(String(lcid.value), 10);

  if (isNaN(lcidValue)) {
    const localeStr = String(lcid.value).toLowerCase();
    const key = findLocaleKey(localeStr);
    if (key) {
      currentLocale = key;
      return { type: 'Long', value: localeToLcid.get(key) ?? 1033 };
    }
    return getLocale();
  }

  const locale = lcidToLocale.get(lcidValue);
  if (locale) {
    currentLocale = locale;
    return { type: 'Long', value: lcidValue };
  }

  return getLocale();
}

export function getCurrentLocaleTag(): string {
  const locale = getBrowserLocale();
  return locale.replace(/-/g, '_').split('_')[0] + (locale.includes('-') ? '-' + locale.split('-')[1].toUpperCase() : '');
}

export function getCurrentBCP47Locale(): string {
  return getBrowserLocale();
}

const localeToCurrency: Map<string, string> = new Map([
  ['en-us', 'USD'],
  ['en-gb', 'GBP'],
  ['en-au', 'AUD'],
  ['en-ca', 'CAD'],
  ['en-nz', 'NZD'],
  ['en-za', 'ZAR'],
  ['de-de', 'EUR'],
  ['de-at', 'EUR'],
  ['de-ch', 'CHF'],
  ['fr-fr', 'EUR'],
  ['fr-be', 'EUR'],
  ['fr-ca', 'CAD'],
  ['fr-ch', 'CHF'],
  ['ja', 'JPY'],
  ['zh-cn', 'CNY'],
  ['zh-tw', 'TWD'],
  ['zh-hk', 'HKD'],
  ['ko', 'KRW'],
  ['it', 'EUR'],
  ['es', 'EUR'],
  ['es-mx', 'MXN'],
  ['pt-br', 'BRL'],
  ['pt-pt', 'EUR'],
  ['ru', 'RUB'],
  ['pl', 'PLN'],
  ['nl', 'EUR'],
  ['sv', 'SEK'],
  ['da', 'DKK'],
  ['no', 'NOK'],
  ['fi', 'EUR'],
  ['cs', 'CZK'],
  ['hu', 'HUF'],
  ['tr', 'TRY'],
  ['th', 'THB'],
  ['id', 'IDR'],
  ['ms', 'MYR'],
  ['vi', 'VND'],
  ['in', 'INR'],
  ['hi', 'INR'],
  ['ar-sa', 'SAR'],
  ['ar-eg', 'EGP'],
  ['ar-ae', 'AED'],
  ['he', 'ILS'],
  ['en-ie', 'EUR'],
  ['en-ph', 'PHP'],
  ['en-in', 'INR'],
]);

export function getCurrentCurrency(): string {
  const locale = getBrowserLocale();
  const currency = localeToCurrency.get(locale);
  if (currency) return currency;
  
  const baseLocale = locale.split('-')[0];
  const baseCurrency = localeToCurrency.get(baseLocale);
  if (baseCurrency) return baseCurrency;
  
  return 'USD';
}

export const localeFunctions = {
  GetLocale: getLocale,
  SetLocale: setLocale,
};
