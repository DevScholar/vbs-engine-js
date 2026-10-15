// The locales whose month names refuse a dot added after them.
//
// Measured on Windows build 26200 25H2, 2026-10-15: for every locale, each month name Windows writes was read
// back as "15 <name> 2026" and again as "15 <name>. 2026". Nearly every locale reads the dotted form as well
// (English "Jan." and "January." are January, German "Januar." is January), and a name that is itself written
// with a dot refuses a second one everywhere: Catalan "gen." reads and "gen.." does not. These locales refuse
// the dot for every name they answer to: 1054 (12/12)
//
// Editing this by hand makes it something other than what Windows answers.

/** The LCIDs that refuse a dot added after a month name. */
export const NLS_NO_TRAILING_DOT: ReadonlySet<number> = new Set([1054]);
