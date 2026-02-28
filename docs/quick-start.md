# Quick Start Guide

This guide will walk you through setting up and running vbs-engine-js examples from scratch.

## Prerequisites

- Node.js (v18 or higher recommended)
- npm

## Step 1: Create Project Directory

```bash
mkdir vbs-engine-js-examples
cd vbs-engine-js-examples
```

## Step 2: Build vbs-engine-js (Dependency)

First, you need to build the vbs-engine-js library since the examples reference it locally:

```bash
# Clone or navigate to vbs-engine-js directory
cd ../vbs-engine-js

# Install dependencies
npm install

# Build the library
npm run build

# Return to examples directory
cd ../vbs-engine-js-examples
```

## Step 3: Browser Example

### 3.1 Create Browser Example Directory

```bash
mkdir browser
cd browser
```

### 3.2 Initialize npm Project

```bash
npm init -y
```

### 3.3 Update package.json

Replace the content of `package.json` with:

```json
{
  "name": "vbs-engine-js-browser-example",
  "version": "1.0.0",
  "type": "module",
  "scripts": {
    "dev": "vite"
  },
  "dependencies": {
    "@devscholar/vbs-engine-js": "^0.0.1"
  },
  "devDependencies": {
    "vite": "^5.0.0"
  }
}
```

### 3.4 Create index.html

Create `index.html` with the following content:

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>VBScript Browser Example</title>
  <script type="module">
    import { VbsEngine } from '@devscholar/vbs-engine-js';
    new VbsEngine({ mode: 'browser' });
  </script>
</head>
<body>
  <h1>VBScript Browser Example</h1>
  <div id="output"></div>

  <script type="text/vbscript">
    message = "Hello from VBScript!"

    Function Add(a, b)
        Add = CLng(a) + CLng(b)
    End Function

    Sub ShowOutput()
        Set el = document.getElementById("output")
        el.innerHTML = message & "<br>" & "2 + 3 = " & Add(2, 3)
    End Sub

    Call ShowOutput()
  </script>
</body>
</html>
```

### 3.5 Install Dependencies and Run

```bash
npm install
npm run dev
```

Open your browser and navigate to `http://localhost:5173/`

You should see:
```
VBScript Browser Example
Hello from VBScript!
2 + 3 = 5
```

## Step 4: Node.js Example

### 4.1 Create Node Example Directory

```bash
cd ..
mkdir node
cd node
```

### 4.2 Initialize npm Project

```bash
npm init -y
```

### 4.3 Update package.json

Replace the content of `package.json` with:

```json
{
  "name": "vbs-engine-js-node-example",
  "version": "1.0.0",
  "type": "module",
  "scripts": {
    "start": "node --experimental-transform-types index.ts"
  },
  "dependencies": {
    "@devscholar/vbs-engine-js": "file:../../vbs-engine-js"
  }
}
```

### 4.4 Create index.ts

Create `index.ts` with the following content:

```typescript
import { VbsEngine } from '@devscholar/vbs-engine-js';

const engine = new VbsEngine();

// Add VBScript functions
engine.addCode(`
  Function Add(a, b)
      Add = a + b
  End Function

  Function Multiply(a, b)
      Multiply = a * b
  End Function
`);

// Call functions
const sum = engine.run('Add', 10, 20);
console.log('10 + 20 =', sum);

const product = engine.run('Multiply', 6, 7);
console.log('6 * 7 =', product);

// Execute statements
engine.executeStatement('x = 100');
engine.executeStatement('y = 200');

// Evaluate expressions
const result = engine.eval('x + y');
console.log('x + y =', result);
```

### 4.5 Install Dependencies and Run

```bash
npm install
npm start
```

Expected output:
```
10 + 20 = 30
6 * 7 = 42
x + y = 300
```

## Project Structure

After completing all steps, your project structure should look like:

```
vbs-engine-js-examples/
├── browser/
│   ├── index.html
│   ├── package.json
│   └── node_modules/
├── node/
│   ├── index.ts
│   ├── package.json
│   └── node_modules/
└── quick-start.md
```

## Troubleshooting

### Error: Cannot resolve dependency

Make sure you have built vbs-engine-js first:
```bash
cd ../vbs-engine-js
npm run build
```

### Error: Module not found

Check that the path in package.json is correct:
- For browser: `"file:../../vbs-engine-js"`
- For node: `"file:../../vbs-engine-js"`

### Error: Cannot find package '@devscholar/vbs-engine-js'

Delete node_modules and package-lock.json, then reinstall:
```bash
rm -rf node_modules package-lock.json
npm install
```

## Next Steps

- Read the [API documentation](../vbs-engine-js/README.md) for more features
- Explore the [examples](../vbs-engine-js/examples/) in the main repository
- Try exposing JavaScript objects to VBScript using `engine.addObject()`
