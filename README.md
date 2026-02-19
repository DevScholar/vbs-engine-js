# VBSEngineJS

⚠️ This project is still in pre-alpha stage. Expect breaking changes.

A VBScript engine implemented in TypeScript, supporting VBScript code execution in browser and Node.js environments.

## Installation

```bash
npm install
```

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
    import { VbsBrowserEngine } from './src/index.ts';
    new VbsBrowserEngine();
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

engine.run(`
  name = "World"
  Function Greet(n)
      Greet = "Hello, " & n & "!"
  End Function
  Print Greet(name)
`);

console.log(engine.getVariableAsJs('name'));
```

Run with `npx tsx examples/node-demo.ts`.


## Browser Runtime Options

The `VbsBrowserEngine` constructor accepts an options object with the following properties:

```typescript
interface BrowserRuntimeOptions {
  parseScriptElement?: boolean;
  parseInlineEventAttributes?: boolean;
  injectGlobalThis?: boolean;
  parseEventSubNames?: boolean;
  maxExecutingTime?: number;
}
```

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `parseScriptElement` | `boolean` | `true` | Automatically process and execute all `<script type="text/vbscript">` tags |
| `parseInlineEventAttributes` | `boolean` | `true` | Automatically process inline event attributes like `onclick="vbscript:..."` |
| `injectGlobalThis` | `boolean` | `true` | Enable IE-style global variable sharing between VBScript and JavaScript |
| `parseEventSubNames` | `boolean` | `true` | Automatically bind event handlers from Sub names like `Button1_OnClick`. Requires `injectGlobalThis: true` |
| `maxExecutingTime` | `number` | `-1` | Maximum script execution time in milliseconds. `-1` means unlimited |

### Example Usage

```typescript
import { VbsBrowserEngine } from './src/index.ts';

new VbsBrowserEngine({
  parseScriptElement: true,
  parseInlineEventAttributes: true,
  injectGlobalThis: true,
  parseEventSubNames: true,
  maxExecutingTime: 5000
});
```

# Compatibility

See [docs/compatibility.md](docs/compatibility.md) for details.