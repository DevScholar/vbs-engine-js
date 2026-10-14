// Development tool: re-derives the claim in the README that FormatDateTime agrees with vbscript.dll in
// every measured cell. It runs every locale and named format the engine knows against an oracle captured
// from the real engine and reports each cell that differs.
//
// The oracle is not kept in the repository. To produce one, run a VBScript under `cscript //E:vbscript` that
// calls SetLocale per locale and FormatDateTime(d, 0..4) for each instant, and have it write its output
// through FileSystemObject.CreateTextFile(path, True, True) - the third argument makes the file UTF-16LE, and
// without it the console's ANSI code page replaces the characters of several locales with "?" before any
// wrapper can see them. One process per locale is the reliable way to measure: SetLocale in a loop reports
// values a fresh process does not.
//
// The file is UTF-16LE and holds one line per locale and instant:
//   <lcid>|<instant>|<style0>~<style1>~<style2>~<style3>~<style4>
// with a refused call written as "ERRN!". The instants, in order, are the ones below.
//
// Usage: node --experimental-transform-types --experimental-vm-modules scripts/verify-oracle.ts <oracle>
import { readFileSync } from 'node:fs';
import { formatVbDateTime } from '../src/runtime/vb-date.ts';
import { LCID_TO_LOCALE } from '../src/builtins/locale-data.generated.ts';

const oraclePath = process.argv[2];
if (!oraclePath) throw new Error('usage: verify-oracle.ts <oracle-file>');
const raw = readFileSync(oraclePath, 'utf16le');

/** The five instants the oracle is captured at, as local wall-clock readings. */
const INSTANTS: [number, number, number, number, number, number][] = [
  [2026, 1, 5, 9, 5, 0],
  [2026, 1, 5, 0, 0, 0],
  [2026, 1, 5, 15, 7, 9],
  [1899, 12, 30, 9, 5, 0],
  [1899, 12, 30, 0, 0, 0],
];

/** Windows writes a refused call as "ERRN!" and the engine reports it as "<err N>"; both mean one thing. */
const errorCode = (text: string): string | undefined =>
  /^<err (\d+)>$/.exec(text)?.[1] ?? /^ERR(\d+)!$/.exec(text)?.[1];

let cells = 0;
let differ = 0;
const locales = new Set<number>();
for (const line of raw.split(/\r?\n/)) {
  if (line === '') continue;
  const [lcidText, instantText, stylesText] = line.split('|');
  const lcid = Number(lcidText);
  const parts = INSTANTS[Number(instantText)];
  if (!parts) continue;
  locales.add(lcid);
  const date = new Date(parts[0], parts[1] - 1, parts[2], parts[3], parts[4], parts[5]);
  const locale = LCID_TO_LOCALE.get(lcid) ?? 'en-us';
  const expected = stylesText.split('~');
  for (let style = 0; style < 5; style++) {
    cells++;
    let got: string;
    try {
      got = formatVbDateTime(date, lcid, style, locale);
    } catch (error) {
      got = `<err ${(error as { number?: number }).number ?? '?'}>`;
    }
    if (got !== expected[style] && errorCode(got) !== errorCode(expected[style])) {
      differ++;
      console.log(`lcid=${lcid} ${locale} instant=${instantText} style=${style}`);
      console.log(`  want=${JSON.stringify(expected[style])}`);
      console.log(`  got =${JSON.stringify(got)}`);
    }
  }
}
console.log(`locales=${locales.size} cells=${cells} differ=${differ}`);
