# VBSEngineJS

A VBScript engine implemented in TypeScript, supporting VBScript code execution in browser and Node.js environments.

## Installation

```bash
npm install
```

## Running Examples

See [quick-start.md](docs/quick-start.md) for more information.

## API Reference

The `VbsEngine` class provides an API similar to Microsoft's MSScriptControl:

### Methods

| Method                                                           | Description                                                                                                                                                                                                                                                                                                                                            |
| ---------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `addCode(code: string)`                                          | Adds script code to the engine (function/class definitions)                                                                                                                                                                                                                                                                                            |
| `executeStatement(statement: string)`                            | Executes a VBScript statement; captures errors in `engine.error`                                                                                                                                                                                                                                                                                       |
| `executeStatementThrows(statement: string)`                      | Executes a statement and rethrows the error as a native JS `Error` with an HRESULT `number` (IE/WSH `catch` style). It also carries `description`, the VBScript `line` and `column` of the statement that failed, and a `stack` whose first line is the reading `cscript` prints, as in `(4, 3) Microsoft VBScript runtime error: Division by zero`    |
| `run(procedureName: string, ...args)`                            | Calls a function and returns the result                                                                                                                                                                                                                                                                                                                |
| `eval(expression: string)`                                       | Evaluates an expression and returns the result                                                                                                                                                                                                                                                                                                         |
| `addObject(name: string, object: unknown, addMembers?: boolean)` | Exposes a JavaScript object to VBScript                                                                                                                                                                                                                                                                                                                |
| `releaseGlobals()`                                               | Ends a script: runs `Class_Terminate` for every class instance the script still holds, the way the end of a `cscript` script does. The engine that claimed host integration calls this by itself when its host ends (the process in Node, the page going away in a browser). A host that runs several scripts through one engine calls it between them |
| `clearError()`                                                   | Clears the last error                                                                                                                                                                                                                                                                                                                                  |
| `destroy()`                                                      | Cleans up resources (browser mode) and takes back the engine's host end hook                                                                                                                                                                                                                                                                           |

### Properties

| Property | Type               | Description                                                                                                                                                                         |
| -------- | ------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `error`  | `VbsError \| null` | The last error that occurred, or null. Mirrors MSScriptControl's `Error` object: `number` (pure code), `description`, `source`, `line`, `column`, `text`, `helpFile`, `helpContext` |

### Example Usage

```typescript
import { VbsEngine } from './src/index.ts';

const engine = new VbsEngine();

// Define functions
engine.addCode(`
  Function Multiply(a, b)
      Multiply = a * b
  End Function
  
  Class Calculator
      Public Function Add(x, y)
          Add = x + y
      End Function
  End Class
`);

// Call functions
const product = engine.run('Multiply', 6, 7);
console.log(product); // 42

// Expose JavaScript objects
const myApp = {
  name: 'MyApp',
  version: '1.0.0',
  greet: (name: string) => `Hello, ${name}!`,
};
engine.addObject('MyApp', myApp, true);

// Use exposed object in VBScript
engine.executeStatement('result = MyApp.greet("World")');
const result = engine.eval('result');
console.log(result); // "Hello, World!"

// Error handling
engine.executeStatement('x = CInt("not a number")');
if (engine.error) {
  console.log('Error:', engine.error.description);
}
```

## COM / ActiveX Integration (Node.js on Windows)

VBScript's `CreateObject` and `GetObject` depend on a host-provided COM backend.
In Node.js, wire up [`@devscholar/node-ps1-dotnet`](https://www.npmjs.com/package/@devscholar/node-ps1-dotnet) before running scripts:

```bash
npm install @devscholar/node-ps1-dotnet
```

```typescript
import { VbsEngine } from '@devscholar/vbs-engine-js';
import { ActiveXObject, GetObject } from '@devscholar/node-ps1-dotnet/activex';

// Expose COM constructors to the engine
globalThis.ActiveXObject = ActiveXObject;
globalThis.GetObject = GetObject;

const engine = new VbsEngine();
```

### `CreateObject`

```typescript
engine.executeStatement(`
  Set fso = CreateObject("Scripting.FileSystemObject")
  MsgBox fso.GetAbsolutePathName(".")
`);
```

### `GetObject` — bind to a running instance

```typescript
// GetObject(, "Excel.Application") — gets a running Excel instance
engine.executeStatement(`
  Set xl = GetObject(, "Excel.Application")
  MsgBox xl.Version
`);
```

### `GetObject` — bind to a WMI moniker

```typescript
engine.executeStatement(`
  Set wmi = GetObject("winmgmts:")
  For Each os In wmi.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    MsgBox os.Caption & " " & os.Version
  Next
`);
```

### `GetObject` — bind to a file

```typescript
engine.executeStatement(`
  Set doc = GetObject("C:\Reports\sales.xls")
  MsgBox doc.Name
`);
```

> **Note:** COM/ActiveX requires Windows. On other platforms `CreateObject` returns a
> placeholder object and `GetObject` returns `Nothing`.

## Engine Options

```typescript
interface VbsEngineOptions {
  mode?: 'general' | 'browser' | 'auto'; // Default: 'auto'
  hostIntegration?: boolean; // Default: true for the first engine in a process, false after
  useJsError?: boolean; // Default: true
  injectGlobalThis?: boolean; // Default: true
  injectVBArrayToGlobalThis?: boolean; // Default: true
  maxExecutionTime?: number; // Default: -1 (unlimited)

  // Browser mode only:
  parseScriptElement?: boolean; // Default: true
  parseInlineEventAttributes?: boolean; // Default: true
  parseEventSubNames?: boolean; // Default: true
  parseVbsProtocol?: boolean; // Default: true
  overrideJsEvalFunctions?: boolean; // Default: true
}
```

| Option                       | Type                     | Default             | Description                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| ---------------------------- | ------------------------ | ------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `mode`                       | `'general' \| 'browser'` | `'auto'`            | Engine mode. Auto-detects based on environment: `'browser'` in browser, `'general'` in Node.js. Use `'browser'` for web applications                                                                                                                                                                                                                                                                                                                                                                                                                             |
| `hostIntegration`            | `boolean`                | first engine `true` | Whether the engine may own the host, as it does in IE or WSH. A host has one of these, so the first engine to ask claims it and every later engine is isolated and cannot ask: it does not own the document, install the `vbscript:` handler, replace the host's `eval` and timers, or publish its functions into the host global. It still reads host globals and still calls objects given to `addObject`. The engine that claims it also lets the script go when the host ends, running `Class_Terminate` there the way `cscript` does. Fixed at construction |
| `useJsError`                 | `boolean`                | `true`              | How an error surfaces when the host runs script no JavaScript call asked for (an inline event attribute, a timer handler). `true` throws a real JavaScript error, which the page can catch; `false` leaves it in `engine.error`. `executeStatement` and `executeStatementThrows` are unaffected: the method called decides                                                                                                                                                                                                                                       |
| `injectGlobalThis`           | `boolean`                | `true`              | Enable IE-style global variable sharing between VBScript and JavaScript. Reads of host globals are unaffected; this only decides whether the engine publishes its own functions and variables, and it is forced off for an isolated engine                                                                                                                                                                                                                                                                                                                       |
| `injectVBArrayToGlobalThis`  | `boolean`                | `true`              | Expose VBArray as `globalThis.VBArray` for legacy code compatibility                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| `maxExecutionTime`           | `number`                 | `-1`                | Maximum script execution time in milliseconds. `-1` means unlimited                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| `parseScriptElement`         | `boolean`                | `true`              | Automatically process `<script type="text/vbscript">` tags (browser mode)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| `parseInlineEventAttributes` | `boolean`                | `true`              | Process inline event attributes like `onclick="vbscript:..."` (browser mode)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     |
| `parseEventSubNames`         | `boolean`                | `true`              | Auto-bind event handlers from Sub names like `Button1_OnClick` (browser mode)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| `parseVbsProtocol`           | `boolean`                | `true`              | Handle `vbscript:` protocol in links (browser mode)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| `overrideJsEvalFunctions`    | `boolean`                | `true`              | Override JS eval functions (`eval`,`setTimeout`,`setInterval`) to support VBScript (browser mode)                                                                                                                                                                                                                                                                                                                                                                                                                                                                |

### Dates and locale

Dates are OLE Automation dates: a day count from 30 December 1899, with the time of day as the fraction.

`CStr` of a date, `CDate` and `IsDate` are measured against `vbscript.dll` rather than guessed at. Two things are worth knowing when reading code that touches dates:

- **Which strings are dates follows the locale**, because that is what `SetLocale` is for: `IsDate("29.09.2026")` is `True` under `1031` and `False` under `1033`, and month names come from the locale (English names are read only for Chinese, Japanese and Korean, whose own names `vbscript.dll` does not match).
- **Which types are dates does not.** `DateValue` and `TimeValue` take a date or a string only: a number, a Boolean or Empty is error 13 there and `Null` is error 94, in every locale. Only `CDate` reads numbers, Booleans and Empty.
- **`CDate` reads a plain number inside a string as a day count, and `IsDate` does not.** `CDate("2026")` is 18 July 1905, the 2026th day after 30 December 1899, while `IsDate("2026")` is `False` and `DateValue("2026")` is error 13. A year under a hundred is read as 2000 to 2049 and 1950 to 1999, by value rather than by the digits written, so `CDate("0001-01-01")` is 1 January 2001 while `CDate("1/1/100")` is the year 100 and prints as `01.01.100`.

`CStr` writes the date from Windows NLS, shipped with the engine as data (`src/runtime/nls-date-table.generated.ts`), because a browser has no NLS to ask and CLDR disagrees with it for most locales. The engine matches `vbscript.dll` for every locale it knows.

`FormatDateTime` reads the same table, so it follows Windows rather than the host: `FormatDateTime(d, 1)` under `1031` is `Montag, 5. Januar 2026` and under `1042` is `2026년 1월 5일 월요일`, neither of which `Intl` produces.

- **Its two date formats use the locale's own calendar, and its other three do not.** `ar-SA` gives the Hijri `16/07/47` for `vbShortDate` while giving the Gregorian `05/01/2026` for `vbGeneralDate`, and `fa-IR` gives the Persian `15 دی 1404` for `vbLongDate`.
- **`vbShortTime` keeps the locale's own separator** with a padded 24-hour hour: `09.05` under `1035`, `09:05` under `1042`.
- **The Arabic and Hebrew date formats mark each field** with a right-to-left mark, and their names use a narrow no-break space between words, as Windows does. `vbGeneralDate` carries neither.
- **Some names come from `src/runtime/name-tables.generated.ts`**, which holds the names Windows writes where CLDR has changed them or has none at all. The Vietnamese `Tháng Giêng`, the Latvian lower-case weekday, the Sotho and Tswana names, Tsonga, Algerian and Tunisian Arabic, Hindi's December and the German and Croatian abbreviations are a few of them.
- **Thai writes its year in the Buddhist calendar in every format**, `CStr` included: `CStr(DateSerial(2026, 1, 5))` is `5/1/2569`. It reads a year the same way, so `CDate("5 มกราคม 2026")` is the Gregorian year 1483, which `CStr` writes back as `5/1/2026`. Its long date carries no weekday, which is what Windows writes there even though the locale's own picture string asks for one.

### Example Usage

```typescript
// General mode (Node.js or browser without auto-parsing)
const engine = new VbsEngine({
  injectGlobalThis: true,
  maxExecutionTime: 5000,
});

// Browser mode (auto-parsing enabled)
const browserEngine = new VbsEngine({
  mode: 'browser',
  parseScriptElement: true,
  parseInlineEventAttributes: true,
  parseEventSubNames: true,
});
```

## Compatibility

See [docs/compatibility.md](docs/compatibility.md) for details.
