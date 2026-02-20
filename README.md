# VBSEngineJS

⚠️ This project is still in pre-alpha stage. Expect breaking changes.

A VBScript engine implemented in TypeScript, supporting VBScript code execution in browser and Node.js environments.

## Installation

```bash
npm install
```

## Running Examples

### HTML Example (Browser)

Start the development server and open the example page:

```bash
npm run dev
```

Then open your browser and visit:
- `http://localhost:5173/examples/index.html` - Full demo with MsgBox, InputBox, clock, and event handlers

Or directly:
```bash
npx vite examples/index.html
```

### Node.js Example

Run the Node.js demo that shows how to use VBScript with Node.js modules:

```bash
npx tsx examples/node-demo.ts
```

This example demonstrates:
- Defining VBScript functions with `addCode()`
- Calling VBScript functions with `run()`
- Accessing Node.js modules (path, fs) from VBScript via globalThis

## Quick Start

### Browser Environment

Write VBScript code using the `<script type="text/vbscript">` or `<script language="vbscript">` tag in HTML:

```html
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>VBScript Demo</title>
  <script type="module">
    import { VbsEngine } from './src/index.ts';
    new VbsEngine({ mode: 'browser' });
  </script>
</head>
<body>
  <div id="output"></div>

  <script type="text/vbscript">
    msg = "Hello from VBScript!"
    
    Function Add(a, b)
        Add = CLng(a) + CLng(b)
    End Function
    
    Sub ShowOutput()
        Set el = document.getElementById("output")
        el.innerHTML = msg & "<br>" & "2 + 3 = " & Add(2, 3)
    End Sub
    
    Call ShowOutput()
  </script>
</body>
</html>
```

Start the development server:

```bash
npm run dev
```

Then visit `http://localhost:5173/examples/index.html`.

### Node.js Environment

```typescript
import { VbsEngine } from './src/index.ts';

const engine = new VbsEngine();

// Add function definitions
engine.addCode(`
  Function Add(a, b)
      Add = a + b
  End Function
`);

// Call a function
const result = engine.run('Add', 10, 20);
console.log('Result:', result);  // 30

// Execute a statement
engine.executeStatement('name = "World"');

// Evaluate an expression
const greeting = engine.eval('"Hello, " & name & "!"');
console.log(greeting);  // "Hello, World!"
```

Run with `npx tsx examples/node-demo.ts`.

## API Reference

The `VbsEngine` class provides an API similar to Microsoft's MSScriptControl:

### Methods

| Method | Description |
|--------|-------------|
| `addCode(code: string)` | Adds script code to the engine (function/class definitions) |
| `executeStatement(statement: string)` | Executes a single VBScript statement |
| `run(procedureName: string, ...args)` | Calls a function and returns the result |
| `eval(expression: string)` | Evaluates an expression and returns the result |
| `addObject(name: string, object: unknown, addMembers?: boolean)` | Exposes a JavaScript object to VBScript |
| `clearError()` | Clears the last error |
| `destroy()` | Cleans up resources (browser mode) |

### Properties

| Property | Type | Description |
|----------|------|-------------|
| `error` | `VbsError \| null` | The last error that occurred, or null |

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
console.log(product);  // 42

// Expose JavaScript objects
const myApp = {
  name: 'MyApp',
  version: '1.0.0',
  greet: (name: string) => `Hello, ${name}!`
};
engine.addObject('MyApp', myApp, true);

// Use exposed object in VBScript
engine.executeStatement('result = MyApp.greet("World")');
const result = engine.eval('result');
console.log(result);  // "Hello, World!"

// Error handling
engine.executeStatement('x = CInt("not a number")');
if (engine.error) {
  console.log('Error:', engine.error.description);
}
```

## Engine Options

```typescript
interface VbsEngineOptions {
  mode?: 'general' | 'browser';  // Default: 'general'
  injectGlobalThis?: boolean;    // Default: true
  maxExecutionTime?: number;     // Default: -1 (unlimited)
  
  // Browser mode only:
  parseScriptElement?: boolean;       // Default: true
  parseInlineEventAttributes?: boolean; // Default: true
  parseEventSubNames?: boolean;       // Default: true
  parseVbsProtocol?: boolean;         // Default: true
  overrideJsEvalFunctions?: boolean;  // Default: true
}
```

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `mode` | `'general' \| 'browser'` | `'general'` | Engine mode. Use `'browser'` for web applications |
| `injectGlobalThis` | `boolean` | `true` | Enable IE-style global variable sharing between VBScript and JavaScript |
| `maxExecutionTime` | `number` | `-1` | Maximum script execution time in milliseconds. `-1` means unlimited |
| `parseScriptElement` | `boolean` | `true` | Automatically process `<script type="text/vbscript">` tags (browser mode) |
| `parseInlineEventAttributes` | `boolean` | `true` | Process inline event attributes like `onclick="vbscript:..."` (browser mode) |
| `parseEventSubNames` | `boolean` | `true` | Auto-bind event handlers from Sub names like `Button1_OnClick` (browser mode) |
| `parseVbsProtocol` | `boolean` | `true` | Handle `vbscript:` protocol in links (browser mode) |
| `overrideJsEvalFunctions` | `boolean` | `true` | Override JS eval functions to support VBScript (browser mode) |

### Example Usage

```typescript
// General mode (Node.js or browser without auto-parsing)
const engine = new VbsEngine({
  injectGlobalThis: true,
  maxExecutionTime: 5000
});

// Browser mode (auto-parsing enabled)
const browserEngine = new VbsEngine({
  mode: 'browser',
  parseScriptElement: true,
  parseInlineEventAttributes: true,
  parseEventSubNames: true
});
```

## Compatibility

See [docs/compatibility.md](docs/compatibility.md) for details.
