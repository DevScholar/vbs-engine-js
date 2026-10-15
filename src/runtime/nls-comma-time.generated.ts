// The locales that read a comma between a date and the time after it.
//
// Measured on Windows build 26200 25H2, 2026-10-15: "15 <month name> 2026, 13:05:09" is read under these five
// and refused under the other 127, and every one of the 132 refuses the same text when the date part is
// numbers ("15 1 2026, 13:05:09"), which is why the comma is allowed only where a month is named.
//
// Editing this by hand makes it something other than what Windows answers.

/** The LCIDs whose date reader takes a comma before the time. */
export const NLS_COMMA_BEFORE_TIME: ReadonlySet<number> = new Set([1038, 1050, 1062, 2074, 3098]);
