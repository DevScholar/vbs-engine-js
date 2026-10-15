// Development tool, not part of the build: writes src/runtime/name-tables.generated.ts from the
// names-<lcid>.txt dumps scripts/dump-locale-names.vbs produces.
//
// Two tables come out of one dump, because a name Windows writes has two forms and they are different
// questions:
//   * NAME_STANDALONE is what MonthName and WeekdayName answer with: the name on its own. Catalan spells
//     January "gener" alone and "de gener" beside a day, and German spells March "Mär" against "März".
//   * NAME_IN_DATE is what a picture string writes, read out of the long date Windows produces.
//
// Only what Intl gets wrong is listed. A locale with no ICU data at all has nothing to compare with, so
// every one of its names is listed: Intl answers in the machine's own language there, and comparing against
// another language's words would store nothing while the runtime went on answering with them.
//
// The name inside a date is found with a pattern built from the locale's own long-date picture, rather than
// with a hand-written regular expression per locale: each name token becomes a capture, each numeric token a
// run of digits, and the text between them is allowed to differ, because Windows writes separators the
// picture does not ask for (the Hijri long date is "رجب/1447" where the picture says "رجب, 1447").
//
// Usage: node scripts/build-name-tables.ts <dump-dir> <out-file> <lcid,lcid,...>
import { readFileSync, writeFileSync } from 'node:fs';
import { NLS_DATE_FORMATS } from '../src/runtime/nls-date-table.generated.ts';
import { hijriLongDatePicture, withoutWeekdayGroup } from '../src/runtime/vb-date.ts';

const [dumpDirectory, outputPath, lcidList] = process.argv.slice(2);
if (!dumpDirectory || !outputPath || !lcidList)
  throw new Error('usage: build-name-tables.ts <dump-dir> <out-file> <lcids>');

/** The Unicode calendar name for a LOCALE_ICALENDARTYPE, as the runtime reads it. */
const CALENDAR_NAME: Record<number, string> = {
  1: 'gregory',
  6: 'islamic-umalqura',
  7: 'buddhist',
  22: 'persian',
  23: 'islamic-umalqura',
};

/** The Intl tag the engine uses for each LCID, read from the generated table so the two cannot drift. */
const localeTable = readFileSync('src/builtins/locale-data.generated.ts', 'utf8');
const tagOf = (lcid: number): string => {
  const found = new RegExp(`\\[${lcid}, '([^']+)'\\]`).exec(localeTable);
  if (!found) throw new Error(`no Intl tag for ${lcid} in locale-data.generated.ts`);
  return found[1];
};

/** Windows separates the fields of an Arabic or Hebrew name with U+00A0; the runtime does the same. */
const scriptName = (value: string, tag: string): string =>
  /^(ar|he)\b/i.test(tag) ? value.replace(/ /g, '\u00a0') : value;

const namePart = (
  formatter: Intl.DateTimeFormat,
  date: Date,
  type: 'month' | 'weekday',
  tag: string
): string =>
  scriptName(formatter.formatToParts(date).find(part => part.type === type)?.value ?? '', tag);

/** What Intl answers when the name stands alone. */
function intlStandalone(tag: string) {
  const monthLong = new Intl.DateTimeFormat(tag, { month: 'long', timeZone: 'UTC' });
  const monthShort = new Intl.DateTimeFormat(tag, { month: 'short', timeZone: 'UTC' });
  const weekdayLong = new Intl.DateTimeFormat(tag, { weekday: 'long', timeZone: 'UTC' });
  const weekdayShort = new Intl.DateTimeFormat(tag, { weekday: 'short', timeZone: 'UTC' });
  const monthsLong: string[] = [];
  const monthsShort: string[] = [];
  const weekdaysLong: string[] = [];
  const weekdaysShort: string[] = [];
  for (let monthNumber = 0; monthNumber < 12; monthNumber += 1) {
    const date = new Date(Date.UTC(2026, monthNumber, 15));
    monthsLong.push(namePart(monthLong, date, 'month', tag));
    monthsShort.push(namePart(monthShort, date, 'month', tag));
  }
  for (let weekdayNumber = 0; weekdayNumber < 7; weekdayNumber += 1) {
    // 2026-01-04 is a Sunday, which is index 0 both here and in WeekdayName.
    const date = new Date(Date.UTC(2026, 0, 4 + weekdayNumber));
    weekdaysLong.push(namePart(weekdayLong, date, 'weekday', tag));
    weekdaysShort.push(namePart(weekdayShort, date, 'weekday', tag));
  }
  return { monthsLong, monthsShort, weekdaysLong, weekdaysShort };
}

/** One dump file, key by key. */
function readDump(lcid: number): Map<string, string> {
  const values = new Map<string, string>();
  const text = readFileSync(`${dumpDirectory}/names-${lcid}.txt`, 'utf16le').replace(/^\uFEFF/, '');
  for (const line of text.split(/\r?\n/)) {
    const separatorAt = line.indexOf('=');
    if (separatorAt > 0) values.set(line.slice(0, separatorAt), line.slice(separatorAt + 1));
  }
  return values;
}

/** The pieces of a picture string: a quoted literal, a run of one letter, or anything else on its own. */
function patternChunks(pattern: string): { text: string; letters: boolean }[] {
  const chunks: { text: string; letters: boolean }[] = [];
  for (let index = 0; index < pattern.length;) {
    const char = pattern[index];
    if (char === "'") {
      const end = pattern.indexOf("'", index + 1);
      const stop = end < 0 ? pattern.length : end + 1;
      chunks.push({ text: pattern.slice(index, stop), letters: false });
      index = stop;
      continue;
    }
    let run = 1;
    while (pattern[index + run] === char) run += 1;
    chunks.push({ text: char.repeat(run), letters: /[A-Za-z]/.test(char) });
    index += run;
  }
  return chunks;
}

const NUMERIC_DATE_TOKEN = /^(d{1,2}|M{1,2}|y{1,4}|h{1,2}|m{1,2}|s{1,2})$/;

const escapeRegExp = (text: string): string => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * A reader for one locale's long date: which name is where, and the pattern that finds it.
 *
 * The picture is transformed the way the runtime transforms it before it is turned into a pattern, so the
 * two cannot disagree about what Windows writes: the Thai weekday group comes out, and the Hijri date gets
 * its separators. What is left is expanded literally, so the pattern only matches the string Windows
 * produces for that picture, and each name token becomes one capture. A right-to-left mark is allowed in
 * front of every field, because Windows writes one there where the runtime's own expansion does too.
 */
function nameReader(
  picture: string,
  calendarType: number
): { kinds: ('month' | 'weekday')[]; regex: RegExp } | null {
  const transformed =
    calendarType === 7
      ? withoutWeekdayGroup(picture)
      : calendarType === 23
        ? hijriLongDatePicture(picture)
        : picture;
  const kinds: ('month' | 'weekday')[] = [];
  let pattern = '^\\u200f?';
  for (const chunk of patternChunks(transformed)) {
    if (chunk.letters && /^M{3,4}$/.test(chunk.text)) {
      kinds.push('month');
      pattern += '\\u200f?(.+?)';
    } else if (chunk.letters && /^d{3,4}$/.test(chunk.text)) {
      kinds.push('weekday');
      pattern += '\\u200f?(.+?)';
    } else if (chunk.letters && NUMERIC_DATE_TOKEN.test(chunk.text)) {
      pattern += '\\u200f?\\d+';
    } else {
      // A quoted literal is the picture's own text with the quotes taken off; anything else is copied.
      const text =
        chunk.text.startsWith("'") && chunk.text.endsWith("'")
          ? chunk.text.slice(1, -1)
          : chunk.text;
      pattern += escapeRegExp(text);
    }
  }
  const regex = new RegExp(`${pattern}$`);
  return kinds.length === 0 ? null : { kinds, regex };
}

const differences = (
  windows: (string | undefined)[],
  intl: (string | undefined)[] | null
): string[] =>
  windows
    .map((value, index) => ({ value, index }))
    .filter(({ value, index }) => value !== undefined && value !== intl?.[index])
    .map(({ value, index }) => `${index}: ${JSON.stringify(value)}`);

const property = (name: string, entries: string[]): string[] =>
  entries.length === 0
    ? []
    : [`    ${name}: {`, ...entries.map(line => `      ${line},`), '    },'];

const inDateBlocks: string[] = [];
const standaloneBlocks: string[] = [];

for (const lcid of lcidList.split(',').map(Number)) {
  const tag = tagOf(lcid);
  const answersInOwnLanguage = Intl.DateTimeFormat.supportedLocalesOf([tag]).length > 0;
  const dump = readDump(lcid);

  // The standalone names.
  const windowsStandalone = {
    monthsLong: [] as string[],
    monthsShort: [] as string[],
    weekdaysLong: [] as string[],
    weekdaysShort: [] as string[],
  };
  for (const [nameKey, targetProperty] of [
    ['nm', 'monthsLong'],
    ['ns', 'monthsShort'],
    ['dw', 'weekdaysLong'],
    ['ds', 'weekdaysShort'],
  ] as const) {
    for (let nameIndex = 1; nameIndex <= (nameKey.startsWith('n') ? 12 : 7); nameIndex += 1) {
      const value = dump.get(`${nameKey}${nameIndex}`);
      if (value !== undefined) windowsStandalone[targetProperty][nameIndex - 1] = value;
    }
  }
  const intlStandaloneNames = answersInOwnLanguage ? intlStandalone(tag) : null;
  const standaloneLines = [
    ...property(
      'monthsLong',
      differences(windowsStandalone.monthsLong, intlStandaloneNames?.monthsLong ?? null)
    ),
    ...property(
      'monthsShort',
      differences(windowsStandalone.monthsShort, intlStandaloneNames?.monthsShort ?? null)
    ),
    ...property(
      'weekdaysLong',
      differences(windowsStandalone.weekdaysLong, intlStandaloneNames?.weekdaysLong ?? null)
    ),
    ...property(
      'weekdaysShort',
      differences(windowsStandalone.weekdaysShort, intlStandaloneNames?.weekdaysShort ?? null)
    ),
  ];
  if (standaloneLines.length > 0) {
    standaloneBlocks.push(
      `  // ${lcid}: ${answersInOwnLanguage ? 'Intl spells these differently from Windows when the name stands alone.' : 'no ICU data, so every name is listed here.'}\n  ${lcid}: {\n${standaloneLines.join('\n')}\n  },`
    );
  }

  // The names inside a date.
  const format = NLS_DATE_FORMATS.get(lcid);
  const calendarType = format?.calendarType ?? 1;
  const reader = format ? nameReader(format.longDate, calendarType) : null;
  if (!reader) continue;
  const calendar = CALENDAR_NAME[calendarType] ?? 'gregory';
  const ownCalendar = calendar !== 'gregory';
  const tagged = `${tag}-u-ca-${calendar}`;
  const monthFormatter = new Intl.DateTimeFormat(`${tagged}`, {
    month: 'long',
    day: 'numeric',
    timeZone: 'UTC',
  });
  const weekdayFormatter = new Intl.DateTimeFormat(tag, { weekday: 'long', timeZone: 'UTC' });
  const calendarMonth = (date: Date): number =>
    Number(
      new Intl.DateTimeFormat(`${tagged}-nu-latn`, { month: 'numeric', timeZone: 'UTC' })
        .formatToParts(date)
        .find(part => part.type === 'month')?.value ?? '0'
    );

  const months: (string | undefined)[] = [];
  const weekdays: (string | undefined)[] = [];
  const monthIntl: (string | undefined)[] = [];
  const weekdayIntl: (string | undefined)[] = [];
  const samples: [string, Date][] = [];
  for (let monthNumber = 1; monthNumber <= 12; monthNumber += 1)
    samples.push([`md${monthNumber}`, new Date(Date.UTC(2026, monthNumber - 1, 15))]);
  for (let weekdayNumber = 1; weekdayNumber <= 7; weekdayNumber += 1)
    samples.push([`wd${weekdayNumber}`, new Date(Date.UTC(2026, 0, 4 + weekdayNumber))]);
  if (ownCalendar) {
    for (let sampleNumber = 1; sampleNumber <= 14; sampleNumber += 1)
      samples.push([`id${sampleNumber}`, new Date(Date.UTC(2026, 0, 5 + (sampleNumber - 1) * 28))]);
  }

  for (const [key, date] of samples) {
    const text = dump.get(key);
    if (text === undefined) continue;
    const found = reader.regex.exec(text);
    if (!found) continue;
    // The month a name belongs to is the calendar's own where the locale has one: the Islamic year is
    // shorter than the Gregorian one, so its months cannot be reached from a Gregorian month number.
    const monthIndex = ownCalendar ? calendarMonth(date) - 1 : date.getUTCMonth();
    const weekdayIndex = date.getUTCDay();
    reader.kinds.forEach((kind, captureIndex) => {
      const value = found[captureIndex + 1];
      if (value === undefined) return;
      if (kind === 'month') {
        months[monthIndex] = value;
        monthIntl[monthIndex] = namePart(monthFormatter, date, 'month', tag);
      } else {
        weekdays[weekdayIndex] = value;
        weekdayIntl[weekdayIndex] = namePart(weekdayFormatter, date, 'weekday', tag);
      }
    });
  }

  const inDateLines = [
    ...property('months', differences(months, answersInOwnLanguage ? monthIntl : null)),
    ...property('weekdays', differences(weekdays, answersInOwnLanguage ? weekdayIntl : null)),
  ];
  if (inDateLines.length > 0) {
    inDateBlocks.push(
      `  // ${lcid}: ${answersInOwnLanguage ? 'Intl writes a different name inside a date.' : 'no ICU data, so every name is listed here.'}\n  ${lcid}: {\n${inDateLines.join('\n')}\n  },`
    );
  }
}

const header = `// Generated by scripts/build-name-tables.ts - do not edit by hand.
//
// The names Windows writes, in the two forms a script can ask for. Kept apart because they are different
// questions: MonthName and WeekdayName answer with the standalone form, and a picture string with the one
// CLDR calls the format form. Catalan spells January "gener" alone and "de gener" beside a day, and German
// spells March "Mär" against "März".
//
// Only what Intl gets wrong is listed. A locale with no ICU data at all has every name listed, because Intl
// answers in the host language there and a comparison would store nothing while the runtime went on
// answering with another language's words.
//
// Keyed by LCID. A month is keyed by a zero-based month and a weekday by the index WeekdayName uses, which
// puts Sunday at 0. A locale whose calendar is not the Gregorian one is keyed by ITS calendar's month, which
// is the month whose name is written; the runtime looks it up the same way.
//
// Regenerating needs the names-<lcid>.txt dumps scripts/dump-locale-names.vbs writes.

/** The names Windows writes inside a formatted date. */
export interface InDateNames {
  /** The full month names, keyed by the locale's own calendar month. */
  months?: Record<number, string>;
  /** The full weekday names, keyed with Sunday at 0. */
  weekdays?: Record<number, string>;
}

/** The names Windows writes when MonthName or WeekdayName asks for one on its own. */
export interface StandaloneNames {
  /** The full month names Windows writes. */
  monthsLong?: Record<number, string>;
  /** The month abbreviations Windows writes. */
  monthsShort?: Record<number, string>;
  /** The full weekday names Windows writes. */
  weekdaysLong?: Record<number, string>;
  /** The weekday abbreviations Windows writes. */
  weekdaysShort?: Record<number, string>;
}

`;

const inDate =
  inDateBlocks.length === 0
    ? 'export const NAME_IN_DATE: Record<number, InDateNames> = {};\n'
    : `export const NAME_IN_DATE: Record<number, InDateNames> = {\n${inDateBlocks.join('\n')}\n};\n`;
const standalone =
  standaloneBlocks.length === 0
    ? 'export const NAME_STANDALONE: Record<number, StandaloneNames> = {};\n'
    : `export const NAME_STANDALONE: Record<number, StandaloneNames> = {\n${standaloneBlocks.join('\n')}\n};\n`;

writeFileSync(outputPath, `${header}${inDate}\n${standalone}`, 'utf8');
console.log(
  `wrote ${outputPath}: ${inDateBlocks.length} locales in date, ${standaloneBlocks.length} standalone`
);
